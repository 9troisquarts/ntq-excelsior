require 'roo'

module NtqExcelsior
  class Importer

    attr_accessor :file, :check, :lines, :options, :status_tracker

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

    def spreadsheet
      return @spreadsheet unless @spreadsheet.nil?

      raise 'File is missing' unless file.present?
  
      @spreadsheet = Roo::Spreadsheet.open(file, self.class.spreadsheet_options || {})
    end

    def required_headers
      return @required_headers if @required_headers

      @required_columns = self.class.schema.select { |field, column_config| !column_config.is_a?(Hash) || !column_config.has_key?(:required) || column_config[:required] }
      @required_line_keys = @required_columns.map{ |k, v| k }
      @required_headers = @required_columns.map{ |k, column_config| column_config.is_a?(Hash) ? column_config[:header] : column_config }.map{|header| header.is_a?(String) ? Regexp.new(header, "i") : header}
      if self.class.primary_key && !@required_line_keys.include?(self.class.primary_key)
        @required_line_keys = @required_line_keys.unshift(self.class.primary_key)
        @required_headers = @required_headers.unshift(Regexp.new(self.class.primary_key.to_s, "i")) 
      end
      @required_headers
    end

    def spreadsheet_data
      spreadsheet.sheet(spreadsheet.sheets[0]).parse(header_search: required_headers)
    end

    def detect_header_scheme
      return @header_scheme if @header_scheme
      @header_scheme = {}
      l = spreadsheet_data[0].dup

      self.class.schema.each do |field, column_config|
        header = column_config.is_a?(Hash) ? column_config[:header] : column_config

        l.each do |parsed_header, _value|
          next unless header.is_a?(Regexp) && parsed_header.match?(header) || header.is_a?(String) && parsed_header == header
          
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

      @lines = spreadsheet_data.map {|line| parse_line(line) }
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

      @record_attributes = self.class.schema.keys.select{|k| k.to_sym != :id && record.respond_to?(:"#{k}=") }
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
      @success = false
      @action = nil
      @errors = []
      
      if (!self.class.autoset.nil? || self.class.autoset)
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
      at = 0
      errors_lines = []
      success_count = 0
      lines.each_with_index do |line, index|
        break if errors_lines.size == self.class.max_error_count 

        result = import_line(line.with_indifferent_access, save: true)
        case result[:status]
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

      { success_count: success_count, errors: errors_lines }
    end

  end
end 