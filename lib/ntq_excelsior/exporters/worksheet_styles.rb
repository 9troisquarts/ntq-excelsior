module NtqExcelsior
  module Exporters
    class WorksheetStyles
      attr_accessor :styles

      DEFAULT_STYLES = {
        date_format: {
          format_code: "dd-mm-yyyy"
        },
        time_format: {
          format_code: "dd-mm-yyyy hh:mm:ss"
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
      }.freeze

      def initialize(styles = {})
        @styles = DEFAULT_STYLES.merge(styles || {})
      end

      def get_style(style_key)
        style = @styles[style_key.to_sym]
        style || {}
      end
    end
  end
end
