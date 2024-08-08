require 'ostruct'

module NtqExcelsior
  class Context < OpenStruct
    attr_accessor :success

    def success?
      @success.nil? || @success
    end

    def error?
      !@success.nil? && !@success
    end
  end
end
