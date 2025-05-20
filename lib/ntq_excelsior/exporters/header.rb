module NtqExcelsior
  module Exporters
    # Handles the generation and management of Excel headers.
    # This class is responsible for creating and formatting header rows in Excel exports.
    # It supports nested headers, column merging, and style application.
    #
    # == Example
    #   columns = [
    #     Column.new(title: "Name", styles: [:bold]),
    #     Column.new(title: "Details", children: [
    #       Column.new(title: "Age"),
    #       Column.new(title: "City")
    #     ])
    #   ]
    #   header = Header.new(columns, worksheet_styles: worksheet_styles)
    #   rows = header.resolve_rows
    #
    class Header
      include CellHelper

      # The columns used to generate headers
      # @return [Array<Column>] array of column definitions
      attr_reader :columns

      # The worksheet styles
      # @return [WorksheetStyles] the worksheet styles
      attr_reader :worksheet_styles

      # The row offset for header placement
      # @return [Integer] number of rows to offset header placement
      attr_reader :offset_row

      # Creates a new Header instance
      #
      # @param columns [Array<Column>] the columns to generate headers from
      # @param offset_row [Integer] the row offset for header placement (default: 0)
      # @example
      #   header = Header.new([Column.new(title: "Name")], { bold: { b: true } })
      def initialize(columns, offset_row = 1, worksheet_styles: nil)
        @columns = columns
        @offset_row = offset_row
        @worksheet_styles = worksheet_styles
      end

      # Resolves header rows configuration
      #
      # @return [Array<Hash>] array of row configurations, each containing :values, :styles, :merge_cells, and :height
      # @example
      #   rows = header.resolve_rows
      #   # => [{ values: ["Name"], styles: [{ b: true }], merge_cells: [], height: nil }]
      def resolve_rows(context: nil)
        return [empty_row] if columns.empty?

        max_depth = columns.map(&:depth).max || 1
        initial_rows = Array.new(max_depth) { empty_row }

        normalize_rows(place_headers(initial_rows, max_depth, context: context))
      end

      private

      # Creates an empty row structure
      #
      # @return [Hash] empty row with initialized arrays for values, styles, and merge_cells
      def empty_row
        {
          values: [],
          styles: [],
          merge_cells: [],
          data_validations: [],
          height: nil
        }
      end

      # Normalizes row lengths to ensure all rows have the same number of columns
      #
      # @param rows [Array<Hash>] array of row configurations to normalize
      # @return [Array<Hash>] normalized rows with equal column counts
      def normalize_rows(rows)
        max_length = rows.map { |row| row[:values].length }.max

        rows.map do |row|
          {
            values: pad_array(row[:values].dup, max_length, nil),
            styles: pad_array(row[:styles].dup, max_length, {}),
            merge_cells: row[:merge_cells].dup,
            height: row[:height],
            data_validations: row[:data_validations].dup
          }
        end
      end

      # Pads an array to a specified length with a default value
      #
      # @param array [Array] the array to pad
      # @param length [Integer] the desired length
      # @param value [Object] the value to use for padding
      # @return [Array] padded array of specified length
      def pad_array(array, length, value)
        return array if array.length >= length

        array + Array.new(length - array.length, value)
      end

      # Places headers in the rows structure
      #
      # @param rows [Array<Hash>] initial row configurations
      # @param depth [Integer] maximum depth of nested headers
      # @return [Array<Hash>] rows with headers placed
      def place_headers(rows, depth, context: nil)
        columns.reduce([0, rows]) do |(current_col, updated_rows), column|
          next [current_col, updated_rows] unless column.visible?(context: context)

          result = place_header(column, current_col, 0, updated_rows, depth)
          [result[:next_col], result[:rows]]
        end.last
      end

      # Places a single header in the rows structure
      #
      # @param column [Column] the column to place
      # @param start_col [Integer] starting column index
      # @param row [Integer] row index
      # @param rows [Array<Hash>] current row configurations
      # @param depth [Integer] maximum depth of nested headers
      # @return [Hash] next column index and updated rows
      def place_header(column, start_col, row, rows, depth)
        if column.has_children?
          place_parent_header(column, start_col, row, rows, depth)
        else
          place_leaf_header(column, start_col, row, rows, depth)
        end
      end

      # Places a parent header with its children
      #
      # @param column [Column] the parent column
      # @param start_col [Integer] starting column index
      # @param row [Integer] row index
      # @param rows [Array<Hash>] current row configurations
      # @param depth [Integer] maximum depth of nested headers
      # @return [Hash] next column index and updated rows
      def place_parent_header(column, start_col, row, rows, depth)
        result = column.children.reduce([start_col, duplicate_rows(rows)]) do |(col, updated_rows), child|
          next [col, updated_rows] unless child.visible?(context: context)

          result = place_header(child, col, row + 1, updated_rows, depth)
          [result[:next_col], result[:rows]]
        end

        child_start_col, rows_with_children = result
        current_width = child_start_col - start_col

        rows_with_cell = place_cell(column, start_col, row, rows_with_children)
        final_rows = if current_width > 1
                       merge_cells(rows_with_cell, row, start_col, current_width, :parent)
                     else
                       rows_with_cell
                     end

        { next_col: child_start_col, rows: final_rows }
      end

      # Places a leaf header (header without children)
      #
      # @param column [Column] the leaf column
      # @param start_col [Integer] starting column index
      # @param row [Integer] row index
      # @param rows [Array<Hash>] current row configurations
      # @param depth [Integer] maximum depth of nested headers
      # @return [Hash] next column index and updated rows
      def place_leaf_header(column, start_col, row, rows, depth)
        rows_with_cell = place_cell(column, start_col, row, rows, validation: true)
        final_rows = row < depth - 1 ? merge_cells(rows_with_cell, row, start_col, depth - row, :leaf) : rows_with_cell

        { next_col: start_col + 1, rows: final_rows }
      end

      # Places a header cell value and its styles
      #
      # @param column [Column] the column containing the header information
      # @param start_col [Integer] column index for the cell
      # @param row [Integer] row index for the cell
      # @param rows [Array<Hash>] current row configurations
      # @return [Array<Hash>] updated rows with the new cell
      def place_cell(column, col, row, rows, validation: false)
        new_rows = duplicate_rows(rows)
        new_rows[row] = empty_row if new_rows[row].nil?
        new_row = ensure_cell_space(new_rows[row], col)

        new_row = new_row.merge(
          values: new_row[:values].dup.tap { |v| v[col] = column.title || "" },
          styles: new_row[:styles].dup.tap { |s| s[col] = get_styles(column.header_styles) }
        )

        if validation && column.list
          # + 1 on row because we want to start from the next, which is not a header
          range = cells_range([col + 1, row + 1], [col + 1, 1_000_000])
          new_row = new_row.merge(
            data_validations: new_row[:data_validations] + [{
              range: range,
              config: column.validations
            }]
          )
        end

        new_rows.dup.tap { |r| r[row] = new_row }
      end

      # Merges cells for a parent header
      #
      # @param rows [Array<Hash>] current row configurations
      # @param row [Integer] row index
      # @param start_col [Integer] starting column index
      # @param width [Integer] number of columns to merge
      # @return [Array<Hash>] updated rows with merge information
      def merge_cells(rows, row, start_col, span, type)
        new_rows = duplicate_rows(rows)
        new_rows[row] = empty_row if new_rows[row].nil?
        new_row = ensure_merge_space(new_rows[row], start_col, span)
        merge_range = case type
                      when :parent
                        cells_range([start_col + 1, offset_row + row], [start_col + span, offset_row + row])
                      when :leaf
                        cells_range([start_col + 1, offset_row + row], [start_col + 1, offset_row + row + span - 1])
                      end

        new_row = new_row.merge(
          merge_cells: new_row[:merge_cells] + [merge_range]
        )

        cleared_rows = clear_merged_cells(new_rows, row, start_col, span, type)
        cleared_rows.dup.tap { |r| r[row] = new_row }
      end

      # Gets the combined styles for a header
      #
      # @param styles [Array<Symbol>] style keys to apply
      # @return [Hash] combined style definitions
      def get_styles(styles)
        styles ||= []
        return {} if styles.empty? || !worksheet_styles

        styles.reduce({}) do |styles_hash, style_key|
          styles_hash.merge(worksheet_styles.get_style(style_key))
        end
      end

      # Ensures an array has at least the specified size
      #
      # @param array [Array] the array to resize
      # @param size [Integer] minimum required size
      # @param default_value [Object] value to use for new elements
      # @return [Array] array with at least the specified size
      def ensure_size(array, size, default_value)
        return array if array.size >= size

        array + Array.new(size - array.size, default_value)
      end

      # Ensures rows array has at least the specified size
      #
      # @param rows [Array<Hash>] current rows array
      # @param min_size [Integer] minimum required size
      # @return [Array<Hash>] array with at least the specified number of rows
      def ensure_rows(rows, min_size)
        return rows if rows.size >= min_size

        rows + Array.new(min_size - rows.size) { empty_row }
      end

      def duplicate_rows(rows)
        rows.map { |r| r&.transform_values(&:dup) || empty_row }
      end

      def ensure_cell_space(row, col)
        base_row = row || empty_row
        {
          values: ensure_size(base_row[:values]&.dup || [], col + 1, nil),
          styles: ensure_size(base_row[:styles]&.dup || [], col + 1, {}),
          merge_cells: base_row[:merge_cells]&.dup || [],
          height: base_row[:height],
          data_validations: base_row[:data_validations]&.dup || []
        }
      end

      def ensure_merge_space(row, start_col, span)
        size = case span
               when Integer then start_col + span
               when Range then span.end
               end

        ensure_cell_space(row, size)
      end

      def clear_merged_cells(rows, row, start_col, span, type)
        new_rows = duplicate_rows(rows)

        case type
        when :parent
          new_row = new_rows[row] || empty_row
          new_values = new_row[:values]&.dup || []
          new_styles = new_row[:styles]&.dup || []

          (1...span).each do |i|
            new_values[start_col + i] = nil
            new_styles[start_col + i] = {}
          end

          new_rows[row] = new_row.merge(values: new_values, styles: new_styles)
        when :leaf
          ((row + 1)...(row + span)).each do |r|
            new_rows[r] = empty_row if new_rows[r].nil?
            new_row = new_rows[r]
            new_values = new_row[:values]&.dup || []
            new_styles = new_row[:styles]&.dup || []

            new_values[start_col] = nil
            new_styles[start_col] = {}

            new_rows[r] = new_row.merge(values: new_values, styles: new_styles)
          end
        end

        new_rows
      end
    end
  end
end
