require 'roo'

module NtqExcelsior
  class Importer

    attr_accessor :file, :check, :lines, :options, :status_tracker

    class << self

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

    def detect_header_scheme(line)
      return @header_scheme if @header_scheme
      @header_scheme = {}
      l = line.dup

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
        header_scheme = detect_header_scheme(line)
        if header.to_s == self.class.primary_key.to_s
          parsed_line[self.class.primary_key] = value
          next
        end

        header_scheme.each do |header, field|
          parsed_line[field.to_sym] = line[header]
        end
      end

      raise Roo::HeaderRowNotFoundError unless (@required_line_keys - parsed_line.keys).size == 0

      parsed_line
    end

    def lines
      return @lines if @lines

      @lines = spreadsheet_data.map {|line| parse_line(line) }
    end

    # id for default query in model
    # line in case an override is needed to find correct record
    def find_or_initialize_record(line)
      raise "Primary key must be set for using the default find_or_initialize" unless self.class.primary_key

      self.class.model_klass.find_or_initialize_by("#{self.class.primary_key}": line[self.class.primary_key.to_sym])
    end

    def import_line(line, save: true)
      record = find_or_initialize_record(line)

      yield(record, line) if block_given?

      status = {}
      return { status: :success } if record.save

      return { status: :error, errors: record.errors.full_messages.join(", ")  }
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