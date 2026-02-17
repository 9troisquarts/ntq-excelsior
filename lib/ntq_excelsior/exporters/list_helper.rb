module NtqExcelsior
  module Exporters
    module ListHelper
      # Creates data validation for a list of values
      #
      # @param list_config [Array, Hash] The list configuration
      # @return [Hash] The data validation configuration
      # @example Simple list
      #   list_data_validation_for_column(["Yes", "No"])
      # @example Complex list
      #   list_data_validation_for_column({
      #     options: ["Yes", "No"],
      #     show_error_message: true,
      #     error: "Invalid value"
      #   })
      def list_data_validation_for_column(list_config)
        return simple_list_validation(list_config) if list_config.is_a?(Array)

        complex_list_validation(list_config)
      end

      # Create a simple data validation when list is an array of values
      #
      # @param options [Array<String>] The list of values
      # @return [Hash] The data validation configuration
      # @example List in schema
      #   list: ["Yes", "No"]
      def simple_list_validation(options)
        {
          type: :list,
          formula1: "\"#{options.join(", ")}\""
        }
      end

      # Creates a complex data validation for a list of values when list is a hash
      #
      # @param config [Hash] The list configuration
      # @return [Hash] The data validation configuration
      # @example List in schema
      #   list: {
      #     options: ["Yes", "No"],
      #     show_error_message: true,
      #     error: "Invalid value"
      #   }
      def complex_list_validation(config)
        validation = {
          type: :list,
          formula1: "\"#{config[:options].join(", ")}\"",
          showErrorMessage: config[:show_error_message] || false,
          showInputMessage: config[:show_input_message] || false
        }

        add_error_message_config(validation, config) if config[:show_error_message]
        add_input_message_config(validation, config) if config[:show_input_message]

        validation
      end

      # Adds an error message configuration to the data validation
      #
      # @param validation [Hash] The data validation configuration
      # @param config [Hash] The list configuration
      # @return [Hash] The updated data validation configuration
      # @example Add error message
      def add_error_message_config(validation, config)
        validation.merge!(
          error: config[:error] || "",
          errorStyle: config[:error_style] || :stop,
          errorTitle: config[:error_title] || ""
        )
      end

      # Adds an input message configuration to the data validation
      #
      # @param validation [Hash] The data validation configuration
      # @param config [Hash] The list configuration
      # @return [Hash] The updated data validation configuration
      # @example Add input message
      def add_input_message_config(validation, config)
        validation.merge!(
          promptTitle: config[:prompt_title] || "",
          prompt: config[:prompt] || ""
        )
      end
    end
  end
end
