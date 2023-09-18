require 'caxlsx'

module NtqExcelsior
  class MultiWorkbookExporter

    attr_accessor :exporters
  
    def initialize(exporters = [])
      @exporters = exporters
    end

    def export
      exports = exporters
      exports = [exporters] if exporters && !exporters.is_a?(Array)

      package = Axlsx::Package.new
      wb = package.workbook
      wb_styles = wb.styles

      exports.each do |exporter|
        exporter.generate_workbook(wb, wb_styles)
      end

      package
    end

  end
end