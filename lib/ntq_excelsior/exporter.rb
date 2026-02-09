# frozen_string_literal: true

require "caxlsx"
require_relative "exporters/column"
require_relative "exporters/cell_helper"
require_relative "exporters/header"
require_relative "exporters/worksheet_styles"
module NtqExcelsior
  # The Exporter class handles the generation of Excel files from data.
  # It provides a flexible way to export data with customizable headers,
  # styles, and data validation.
  #
  # @example Basic usage
  #   data = [{ name: "John", age: 30 }, { name: "Jane", age: 25 }]
  #   exporter = Exporter.new(data)
  #   exporter.export
  #
  # @attr_accessor [Array] data The data to be exported
  # @attr_accessor [Object] context Additional context for the export
  # @attr_accessor [Proc] progression_tracker A proc to track export progress
  class Exporter
    include Exporters::CellHelper

    attr_accessor :data, :context, :progression_tracker

    class << self
      # Sets or gets the schema for the export
      #
      # @param value [Hash, Proc] The schema configuration or a proc that returns it
      # @return [Hash, Proc] The current schema
      def schema(value = nil)
        return @schema if value.nil? && defined?(@schema)

        @schema ||= value
      end

      # Sets or gets the styles for the export
      #
      # @param value [Hash] The styles configuration
      # @return [Hash] The current styles
      def styles(value = nil)
        @styles ||= value
      end
    end

    # Initializes a new Exporter instance
    #
    # @param data [Array] The data to be exported
    # @example
    #   exporter = Exporter.new([{ name: "John", age: 30 }])
    def initialize(data)
      @data = data
      @data_count = data.size.to_d
      @worksheet_styles = worksheet_styles
    end

    # Returns the schema for the export
    #
    # @return [Hash] The schema configuration
    # @note If the schema is a Proc, it will be called with the context and data
    def schema
      raw_schema = self.class.schema.is_a?(Proc) ? self.class.schema.call(context, data) : self.class.schema
      raw_schema.merge(columns: columns)
    end

    # Returns the columns for the export
    #
    # @return [Array<Column>] The column configurations
    def columns
      return @columns if defined?(@columns)

      raw_schema = self.class.schema.is_a?(Proc) ? self.class.schema.call(context, data) : self.class.schema
      @columns = raw_schema[:columns].map { |col| Exporters::Column.new(col, worksheet_styles: worksheet_styles) }
    end

    # Returns the header manager
    #
    # @return [Header] The header manager instance
    def header
      offset_row = extra_headers.size + 1
      @header ||= Exporters::Header.new(columns, offset_row, worksheet_styles: worksheet_styles.dup)
    end

    def extra_headers
      return [] unless schema[:extra_headers].present?

      extra_heads = schema[:extra_headers]
      extra_heads = [extra_heads] if extra_heads.is_a?(Hash)
      heads = []
      extra_heads.each_with_index do |header_line, index|
        columns = header_line.map { |col| Exporters::Column.new(col, worksheet_styles: worksheet_styles.dup) }
        heads << NtqExcelsior::Exporters::Header.new(columns, index + 1, worksheet_styles: worksheet_styles.dup)
      end
      heads
    end

    # Returns the styles configuration
    #
    # @return [Hash] The styles configuration
    def styles
      self.class.styles
    end

    def worksheet_styles
      return @worksheet_styles if defined?(@worksheet_styles)

      @worksheet_styles = Exporters::WorksheetStyles.new(styles)
    end

    # Exports the data to an Excel file
    #
    # @return [Axlsx::Package] The Excel package ready to be saved
    # @example
    #   exporter = Exporter.new(data)
    #   package = exporter.export
    #   package.serialize("output.xlsx")
    def export
      package = Axlsx::Package.new
      wb = package.workbook
      wb_styles = wb.styles

      generate_workbook(wb, wb_styles)

      package
    end

    def lines
      content[:rows]
    end

    def generate_workbook(workbook, wb_styles)
      workbook.add_worksheet(name: schema[:name]) do |sheet|
        add_sheet_content content, wb_styles, sheet
      end
    end

    private

    # Resolves a header row configuration
    #
    # @param headers [Array<Hash>] The headers configuration
    # @param index [Integer] The current row index
    # @return [Array<Hash>] The resolved row configuration
    def resolve_header_row(header)
      return [{ values: [], styles: [], merge_cells: [], height: nil }] unless header

      header.resolve_rows(context: context)
    end

    # Extracts a nested value from an object using dot notation
    #
    # @param value [Object] The source object
    # @param accessors [Array<String>] The path to the desired value
    # @return [Object] The extracted value
    def dig_value(value, accessors = [])
      v = value
      return v if accessors.empty?

      return v.dig(*accessors) if v.is_a?(Hash)

      v = v.send(accessors[0])
      return v if accessors.length == 1

      dig_value(v, accessors[1..-1])
    end

    # Resolves a record row configuration
    #
    # @param schema [Array<Hash>] The schema configuration
    # @param record [Hash] The current record
    # @param index [Integer] The current row index
    # @return [Hash] The resolved row configuration
    def resolve_record_row(_schema, record, index)
      row = { values: [], styles: [], merge_cells: [], height: nil, types: [] }
      col_index = 1

      columns.each do |column|
        result = column.process(record, index, col_index, context: context)
        col_index = result[:next_col_index]

        row[:values].concat(result[:values])
        row[:types].concat(result[:types])
        row[:styles].concat(result[:styles])
        row[:merge_cells].concat(result[:merge_cells])
      end

      row
    end

    # Generates the content for the Excel sheet
    #
    # @return [Hash] The sheet content with rows and styles
    def content
      content = { rows: [] }
      extra_headers.each do |header|
        content[:rows].concat resolve_header_row(header)
      end
      headers = header.resolve_rows(context: context)
      content[:rows].concat(headers)
      @data.each_with_index do |record, data_index|
        current_index = data_index + 1
        if progression_tracker.is_a?(Proc)
          at = (((current_index.to_d / @data_count) * 100.to_d) / 2).round(2)
          progression_tracker.call(at) if (at % 5).zero?
        end
        row_content = resolve_record_row(columns, record, current_index)
        next unless row_content

        if row_content.is_a?(Array)
          content[:rows].concat(row_content)
        else
          content[:rows] << row_content
        end
      end
      content
    end

    # Adds content to an Excel worksheet
    #
    # @param content [Hash] The content to add
    # @param wb_styles [Axlsx::Styles] The workbook styles
    # @param sheet [Axlsx::Worksheet] The worksheet
    # @return [Axlsx::Worksheet] The modified worksheet
    def add_sheet_content(content, wb_styles, sheet)
      content[:rows].each_with_index do |row, index|
        row_style = []
        if row[:styles].is_a?(Array) && row[:styles].any?
          row[:styles].each do |style|
            row_style << wb_styles.add_style(style || {})
          end
        end
        sheet.add_row row[:values], style: row_style, height: row[:height], types: row[:types]
        if progression_tracker.is_a?(Proc)
          at = 50 + ((((index + 1).to_d / @data_count) * 100.to_d) / 2).round(2)
          progression_tracker.call(at) if (at % 5).zero? || index == content[:rows].length - 1
        end

        row[:data_validations]&.each do |validation|
          sheet.add_data_validation(validation[:range], validation[:config])
        end

        row[:merge_cells]&.each do |range|
          sheet.merge_cells range
        end
      end

      return sheet unless content[:rows].present?

      content[:styles]&.each do |(range, sty)|
        next unless range && sty

        sheet.add_style range, sty.except(:border)
        sheet.add_border range, sty[:border] if sty[:border]
      rescue NoMethodError
        next
      end

      sheet.column_widths * content[:col_widths] if content[:col_widths].present?
      sheet
    end
  end
end
