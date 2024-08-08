require 'roo'
require 'ntq_excelsior/context'

module NtqExcelsior
  class Importer
    attr_accessor :file, :check, :lines, :options, :status_tracker, :success
    attr_reader :context

    class << self
      def autosave(value = nil)
        @autosave ||= value
      end

      def autoset(value = nil)
        @autoset ||= value
      end

      def spreadsheet_options(value = nil)
        @spreadsheet_options ||= value
      end

      def primary_key(value = nil)
        @primary_key ||= value
      end

      def model_klass(value = nil)
        @model_klass ||= value
      end

      def schema(value = nil)
        @schema ||= value
      end

      def before(&block)
        @before = block if block_given?
        @before
      end

      def after(&block)
        @after = block if block_given?
        @after
      end

      def max_error_count(value = nil)
        @max_error_count ||= value
      end

      def structure(value = nil)
        @structure ||= value
      end

      def sample_file(value = nil)
        @sample_file ||= value
      end
    end

    def initialize
      @context = NtqExcelsior::Context.new
    end

    def spreadsheet
      return @spreadsheet unless @spreadsheet.nil?

      raise 'File is missing' unless file.present?

      @spreadsheet = Roo::Spreadsheet.open(file, self.class.spreadsheet_options || {})
    end

    def required_headers
      return @required_headers if @required_headers

      @required_columns = self.class.schema.select { |_field, column_config| !column_config.is_a?(Hash) || !column_config.key?(:required) || column_config[:required] }
      @required_headers = @required_columns.values.map { |column| get_column_header(column) }.map { |header| transform_header_to_regexp(header) }
      if self.class.primary_key && !@required_columns.keys.include?(self.class.primary_key)
        @required_headers.unshift(Regexp.new(self.class.primary_key.to_s, "i"))
      end
      @required_headers
    end

    def spreadsheet_data
      begin
        spreadsheet_data = spreadsheet.sheet(spreadsheet.sheets[0]).parse(header_search: required_headers)
        raise 'File is inconsistent, please check you have data in it or check for invalid characters in headers like , / ; etc...' unless spreadsheet_data.size > 0

        spreadsheet_data
      rescue Roo::HeaderRowNotFoundError => e
        missing_headers = []

        e.message.delete_prefix('[').delete_suffix(']').split(",").map(&:strip).each do |header_missing|
          header_missing_regex = transform_header_to_regexp(header_missing, true)
          header_found = @required_columns.values.find do |column|
            transform_header_to_regexp(get_column_header(column)) == header_missing_regex
          end
          if header_found && header_found.is_a?(Hash)
            if header_found[:header].is_a?(String)
              missing_headers << header_found[:header]
            else
              missing_headers << (header_found[:humanized_header] || header_missing)
            end
          elsif header_found&.is_a?(String)
            missing_headers << header_found
          else
            missing_headers << header_missing
          end
        end
        raise Roo::HeaderRowNotFoundError, missing_headers.join(", ")
      end
    end

    def detect_header_scheme
      return @header_scheme if @header_scheme

      @header_scheme = {}
      # Read the first line of file (not header)
      l = spreadsheet_data[0].dup || []

      self.class.schema.each do |field, column_config|
        header = column_config.is_a?(Hash) ? column_config[:header] : column_config

        l.each do |parsed_header, _value|
          next unless parsed_header
          next unless (header.is_a?(Regexp) && parsed_header && parsed_header.match?(header)) || header.is_a?(String) && parsed_header == header

          l.delete(parsed_header)
          @header_scheme[parsed_header] = field
        end
      end
      @header_scheme[self.class.primary_key.to_s] = self.class.primary_key.to_s if self.class.primary_key && !self.class.schema[self.class.primary_key.to_sym]

      @header_scheme
    end

    def parse_line(line)
      parsed_line = {}
      line.each do |header, value|
        header_scheme = detect_header_scheme
        if header.to_s == self.class.primary_key.to_s
          parsed_line[self.class.primary_key] = value
          next
        end

        header_scheme.each do |header, field|
          parsed_line[field.to_sym] = line[header]
        end
      end

      parsed_line
    end

    def lines
      return @lines if @lines

      @lines = spreadsheet_data.map { |line| parse_line(line) }
    end

    # id for default query in model
    # line in case an override is needed to find correct record
    def find_or_initialize_record(line)
      return nil unless self.class.primary_key && self.class.model_klass

      if line[self.class.primary_key.to_sym].present?
        if self.class.primary_key.to_sym == :id
          record = self.class.model_klass.constantize.find_by id: line[self.class.primary_key.to_sym]
        else
          record = self.class.model_klass.constantize.find_or_initialize_by("#{self.class.primary_key}": line[self.class.primary_key.to_sym])
        end
      end
      record = self.class.model_klass.constantize.new unless record
      record
    end

    def record_attributes(record)
      return @record_attributes if @record_attributes

      @record_attributes = self.class.schema.keys.select { |k| k.to_sym != :id && record.respond_to?(:"#{k}=") }
    end

    def set_record_fields(record, line)
      attributes_to_set = record_attributes(record)
      attributes_to_set.each do |attribute|
        record.send(:"#{attribute}=", line[attribute])
      end
      record
    end

    def import_line(line, save: true)
      record = find_or_initialize_record(line)
      return { status: :not_found } unless record

      @success = false
      @action = nil
      @errors = []

      if (self.class.autoset)
        record = set_record_fields(record, line)
      end

      yield(record, line) if block_given?

      if (self.class.autosave.nil? || self.class.autosave)
        @action = record.persisted? ? 'update' : 'create'
        if save
          @success = record.save
        else
          @success = record.valid?
        end
        @errors = record.errors.full_messages.concat(@errors) if record.errors.any?
      end

      return { status: :success, action: @action } if @success

      return { status: :error, errors: @errors.join(", ") }
    end

    def import(save: true, status_tracker: nil)
      self.class.before.call(@context, options) if self.class.before.is_a?(Proc)
      at = 0
      errors_lines = []
      success_count = 0
      not_found_count = 0
      lines.each_with_index do |line, index|
        break if errors_lines.size == self.class.max_error_count

        result = import_line(line.with_indifferent_access, save: true)
        case result[:status]
        when :not_found
          not_found_count += 1
        when :success
          success_count += 1
        when :error
          error_line = line.map { |k, v| v }
          error_line << result[:errors]
          errors_lines.push(error_line) 
        end

        if @status_tracker&.is_a?(Proc)
          at = (((index + 1).to_d / lines.size) * 100.to_d) 
          @status_tracker.call(at)
        end
      end

      import_stats = { success_count: success_count, not_found_count: not_found_count, errors: errors_lines }
      @context.success = true if errors_lines.empty?
      self.class.after.call(@context, options) if self.class.after.is_a?(Proc)
      import_stats
    end
    
  private
  
    def get_column_header(column)
      return column unless column.is_a?(Hash)
  
      column[:header]
    end
  
    def transform_header_to_regexp(header, gsub_enclosure = false)
      return header unless header.is_a?(String)

      if gsub_enclosure && header.scan(/^\/[\^]?([^(\$\/)]+)[\$]?\/[i]?$/i) && $1
        header = $1
      end
      Regexp.new("^#{header}$", "i")
    end

  end

end