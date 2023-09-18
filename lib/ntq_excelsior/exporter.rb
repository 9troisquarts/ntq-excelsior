require 'caxlsx'

module NtqExcelsior
  class Exporter
    attr_accessor :data
    attr_accessor :context

    DEFAULT_STYLES = {
      date_format: {
        format_code: 'dd-mm-yyyy'
      },
      time_format: {
        format_code: 'dd-mm-yyyy hh:mm:ss'
      },
      bold: {
        b: true
      },
      italic: {
        i: true
      },
      center: {
        alignment: { wrap_text: true }
      }
    }

    COLUMN_NAMES = Array('A'..'Z').freeze

    class << self
      def schema(value = nil)
        @schema ||= value
      end
      def styles(value = nil)
        @styles ||= value
      end
    end

		def initialize(data)
      @data = data
    end

    def schema
      self.class.schema
    end

    def styles
      self.class.styles
    end

    def column_name(col_index)
      index = col_index - 1
      return COLUMN_NAMES[index] if index < 26

      letters = []
      letters << index % 26

      while index >= 26 do
        index = (index / 26) - 1
        letters << index % 26
      end

      letters.reverse.map { |i| COLUMN_NAMES[i] }.join
    end

    def cell_name(col, row, *lock)
      "#{lock.include?(:col) ? '$' : ''}#{column_name(col)}#{lock.include?(:row) ? '$' : ''}#{row}"
    end

    def cells_range(starting = [], ending = [])
      "#{cell_name(*starting)}:#{cell_name(*ending)}"
    end

    def number_of_headers_row(columns, count = 1)
      columns_with_children = columns.select{ |c| c[:children] && c[:children].any? }
      return count unless columns_with_children && columns_with_children.size > 0

      columns_with_children.each do |column|
        number_of_children = number_of_headers_row(column[:children], count += 1) 
        count = number_of_children if number_of_children > count
      end
      count
    end

    def get_styles(row_styles, cell_styles = [])
      row_styles ||= []
      return {} if row_styles.length == 0 && cell_styles.length == 0

      styles_hash = {}
      stylesheet = styles || {}
      (row_styles + cell_styles).each do |style_key|
        styles_hash = styles_hash.merge(stylesheet[style_key] || DEFAULT_STYLES[style_key] || {})
      end
      styles_hash
    end

    def column_is_visible?(column, record = nil)
      return true if !column.key?(:visible)
      return column[:visible].call(record, context) if column[:visible].is_a?(Proc)

      column[:visible]
    end

    def column_width(column)
      return column[:width].call(context) if column[:width] && column[:width].is_a?(Proc)

      column[:width] || 1
    end

    def resolve_header_row(headers, index)
      row = { values: [], styles: [], merge_cells: [], height: nil }
      return row unless headers

      col_index = 1
      headers.each do |header|
        next unless column_is_visible?(header)

        width = column_width(header)
        row[:values] << header[:title] || ''
        row[:styles] << get_styles(header[:header_styles] || header[:styles])
        if width > 1
          colspan = width - 1
          row[:values].push(*Array.new(colspan, nil))
          row[:merge_cells].push cells_range([col_index, index], [col_index + colspan, index])
          col_index += colspan
        end

        col_index += 1
      end
      row
    end

    def dig_value(value, accessors = [])
      v = value
      return  v unless accessors && accessors.length > 0

      return v.dig(*accessors) if v.is_a?(Hash)

      v = v.send(accessors[0])
      return v if accessors.length == 1
      return dig_value(v, accessors.slice(1..-1))
    end

    def format_value(resolver, record)
      styles = []
      type = nil
      if resolver.is_a?(Proc)
        value = resolver.call(record) 
      else
        accessors = resolver
        accessors = accessors.split(".") if accessors.is_a?(String)
        value = dig_value(record, accessors)
      end
      if value.is_a?(String)
        type = :string
      end
      if value.is_a?(Date)
        value = value.strftime("%Y-%m-%d")
        styles << :date_format
        type = :date
      end
      if value.is_a?(Time) | value.is_a?(DateTime)
        value = value.strftime("%Y-%m-%d %H:%M:%S")
        styles << :time_format
        type = :time
      end
      { value: value, styles: styles, type: type }
    end

    def resolve_record_row(schema, record, index)
      row = { values: [], styles: [], merge_cells: [], height: nil, types: [] }
      col_index = 1
      schema.each do |column|
        next unless column_is_visible?(column, record)

        width = column_width(column)
        formatted_value = format_value(column[:resolve], record)
        row[:values] << formatted_value[:value]
        row[:types] << (column[:type] || formatted_value[:type])
        row[:styles] << get_styles(column[:styles], formatted_value[:styles])
        if width > 1
          colspan = width - 1
          row[:values].push(*Array.new(colspan, nil))
          row[:merge_cells].push cells_range([col_index, index], [col_index + colspan, index])
          col_index += colspan
        end

        col_index += 1
      end
      row
    end

    def content
      content = { rows: [] }
      index = 0
      (schema[:extra_headers] || []).each_with_index do |header|
        index += 1
        content[:rows] << resolve_header_row(header, index)
      end
      index += 1
      content[:rows] << resolve_header_row(schema[:columns], index)
      @data.each do |record|
        index += 1
        content[:rows] << resolve_record_row(schema[:columns], record, index)
      end
      content
    end

    def add_sheet_content(content, wb_styles, sheet)
      content[:rows].each do |row|
        row_style = []
        if row[:styles].is_a?(Array) && row[:styles].any?
          row[:styles].each do |style|
            row_style << wb_styles.add_style(style || {})
          end
        end
        sheet.add_row row[:values], style: row_style, height: row[:height], types: row[:types]
        if row[:merge_cells]
          row[:merge_cells]&.each do |range|
            sheet.merge_cells range
          end
        end
      end

      # do not apply styles if there are no rows
      if content[:rows].present?
        content[:styles]&.each_with_index do |(range, sty), index|
          begin
            sheet.add_style range, sty.except(:border) if range && sty
            sheet.add_border range, sty[:border] if range && sty && sty[:border]
          rescue NoMethodError
            # do not apply styles if error
          end
        end

        sheet.column_widths * content[:col_widths] if content[:col_widths].present?
      end

      sheet
    end

    def generate_workbook(wb, wb_styles)
      columns = schema[:columns]
      wb.add_worksheet(name: schema[:name]) do |sheet|
        add_sheet_content content, wb_styles, sheet
      end
    end

		def export
      package = Axlsx::Package.new
      wb = package.workbook
      wb_styles = wb.styles

      generate_workbook(wb, wb_styles)

      package
    end

	end
end