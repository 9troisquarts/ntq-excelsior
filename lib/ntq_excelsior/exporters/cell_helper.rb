module NtqExcelsior
  module Exporters
    # Helper module for Excel cell operations
    # Provides methods for cell name and range calculations
    module CellHelper
      # Array of Excel column names from A to Z
      COLUMN_NAMES = Array("A".."Z").freeze

      # Converts a column index to an Excel column name (A, B, C, ..., AA, AB, etc.)
      #
      # @param col_index [Integer] The 1-based column index
      # @return [String] The Excel column name
      # @example
      #   column_name(1) # => "A"
      #   column_name(27) # => "AA"
      def column_name(col_index)
        index = col_index - 1
        return COLUMN_NAMES[index] if index < 26

        letters = []
        letters << index % 26

        while index >= 26
          index = (index / 26) - 1
          letters << index % 26
        end

        letters.reverse.map { |i| COLUMN_NAMES[i] }.join
      end

      # Generates a cell reference with optional row/column locking
      #
      # @param col [Integer] The column number
      # @param row [Integer, nil] The row number
      # @param lock [Array<Symbol>] Lock options (:col, :row)
      # @return [String] The cell reference
      # @example
      #   cell_name(1, 1) # => "A1"
      #   cell_name(1, 1, :col) # => "$A1"
      def cell_name(col, row = nil, *lock)
        "#{lock.include?(:col) ? "$" : ""}#{column_name(col)}#{lock.include?(:row) ? "$" : ""}#{row}"
      end

      # Generates a range reference between two cells
      #
      # @param starting [Array] The starting cell coordinates [col, row]
      # @param ending [Array] The ending cell coordinates [col, row]
      # @return [String] The range reference
      # @example
      #   cells_range([1, 1], [3, 3]) # => "A1:C3"
      def cells_range(starting = [], ending = [])
        "#{cell_name(*starting)}:#{cell_name(*ending)}"
      end
    end
  end
end
