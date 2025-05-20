# frozen_string_literal: true

require_relative "cell_helper"

module NtqExcelsior
  module Exporters
    # Represents a column in the Excel export
    # Handles column definition, visibility, width, and value resolution
    class Column
      include CellHelper

      attr_reader :title, :resolve, :styles, :type, :width, :visible, :children, :header_styles

      attr_accessor :worksheet_styles

      # Initializes a new Column instance
      #
      # @param config [Hash] The column configuration
      # @option config [String] :title The column header title
      # @option config [Proc, String] :resolve The value resolver
      # @option config [Array<Symbol>] :styles The column styles
      # @option config [Symbol] :type The column data type
      # @option config [Integer, Proc] :width The column width
      # @option config [Boolean, Proc] :visible The column visibility
      # @option config [Array<Hash>] :children Child columns for nested headers
      # @option config [Array<Symbol>] :header_styles Styles specific to the header
      def initialize(config, worksheet_styles: nil)
        @title = config[:title]
        @resolve = config[:resolve]
        @styles = config[:styles] || []
        @type = config[:type]
        @width = config[:width] || 1
        @visible = config.key?(:visible) ? config[:visible] : true
        @children = config[:children]&.map { |child| self.class.new(child, worksheet_styles: worksheet_styles) }
        @header_styles = config[:header_styles] || config[:styles] || []
        @worksheet_styles = worksheet_styles
      end

      # Checks if the column should be visible
      #
      # @param record [Hash, nil] The current record being processed
      # @param context [Object] Additional context for visibility determination
      # @return [Boolean] Whether the column should be visible
      def visible?(record = nil, context: nil)
        return true unless @visible.is_a?(Proc)

        @visible.call(record, context)
      end

      def hidden?
        !visible?
      end

      # Gets the effective width of the column
      #
      # @param context [Object] Additional context for width calculation
      # @return [Integer] The column width
      def effective_width(context = nil)
        return @width.call(context) if @width.is_a?(Proc)

        @width
      end

      # Formats a value according to its type and returns style information
      #
      # @param record [Hash] The data record
      # @return [Hash] The formatted value with style information
      def format_value(record)
        styles = []
        type = nil
        value = resolve_value(record) || ""
        type = :string if value.is_a?(String)
        if value.is_a?(Date)
          value = value.strftime("%Y-%m-%d")
          styles << :date_format
          type = :date
        end
        if value.is_a?(Time) || value.is_a?(DateTime)
          value = value.strftime("%Y-%m-%d %H:%M:%S")
          styles << :time_format
          type = :time
        end

        { value: value, styles: styles, type: type }
      end

      # Checks if the column has children
      #
      # @return [Boolean] Whether the column has child columns
      def has_children?
        @children&.any?
      end

      def depth
        @children&.map(&:depth)&.max || 1
      end

      # Processes the column and returns its computed values
      #
      # @param record [Hash] The current record
      # @param index [Integer] The current row index
      # @param col_index [Integer] The current column index
      # @param context [Object] Additional context
      # @return [Hash] The processed column data and next column index
      def process(record, index, col_index, context: nil)
        return empty_result(col_index) unless visible?(record, context: context)

        return process_children(record, index, col_index, context) if has_children?

        process_single(record, index, col_index, context)
      end

      private

      # Returns an empty result with the next column index
      #
      # @param col_index [Integer] The current column index
      # @return [Hash] Empty result structure
      def empty_result(col_index)
        {
          values: [],
          types: [],
          styles: [],
          merge_cells: [],
          next_col_index: col_index
        }
      end

      # Processes children columns
      #
      # @param record [Hash] The current record
      # @param index [Integer] The current row index
      # @param col_index [Integer] The current column index
      # @param context [Object] Additional context
      # @return [Hash] The processed children columns data
      def process_children(record, index, col_index, context)
        result = empty_result(col_index)

        children.each do |child|
          child_result = child.process(record, index, result[:next_col_index], context: context)

          result[:values].concat(child_result[:values])
          result[:types].concat(child_result[:types])
          result[:styles].concat(child_result[:styles])
          result[:merge_cells].concat(child_result[:merge_cells])
          result[:next_col_index] = child_result[:next_col_index]
        end

        result
      end

      # Processes a single column
      #
      # @param record [Hash] The current record
      # @param index [Integer] The current row index
      # @param col_index [Integer] The current column index
      # @param context [Object] Additional context
      # @return [Hash] The processed single column data
      def process_single(record, index, col_index, context)
        width = effective_width(context)
        formatted_value = format_value(record)
        result = {
          values: [formatted_value[:value]],
          types: [@type || formatted_value[:type]],
          styles: [get_styles(@styles, formatted_value[:styles])],
          merge_cells: [],
          next_col_index: col_index + 1
        }
        if width > 1
          colspan = width - 1
          result[:values].concat!(Array.new(colspan, nil))
          result[:merge_cells] << cells_range([col_index, index], [col_index + colspan, index])
          result[:next_col_index] += colspan
        end

        result
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

      # Resolves the value for the column from the record
      #
      # @param record [Hash] The data record
      # @return [Object] The resolved value
      def resolve_value(record)
        if @resolve.is_a?(Proc)
          @resolve.call(record)
        else
          accessors = @resolve
          accessors = accessors.split(".") if accessors.is_a?(String)
          dig_value(record, accessors)
        end
      end

      # Gets the combined styles for a cell
      #
      # @param row_styles [Array<Symbol>] The row-level styles
      # @param cell_styles [Array<Symbol>] The cell-level styles
      # @return [Hash] The combined styles
      def get_styles(row_styles, cell_styles = [])
        row_styles ||= []
        return {} if (row_styles.empty? && cell_styles.empty?) || !worksheet_styles

        styles_hash = {}
        (row_styles + cell_styles).each do |style_key|
          styles_hash = styles_hash.merge(worksheet_styles.get_style(style_key))
        end
        styles_hash
      end
    end
  end
end
