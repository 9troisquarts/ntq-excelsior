require 'spec_helper'
require_relative '../../samples/user_importer'

RSpec.describe NtqExcelsior::Importer do

  let(:file_path) { File.expand_path('../fixtures/users.xlsx', __dir__) }
  
  let(:importer) do
    user_importer = UserImporter.new
    user_importer.file = file_path
    user_importer
  end

  describe "lines" do
    it "should return the number of lines in the file" do
      expect(importer.lines.size).to eq(1)
    end

    it "should have a key first_name filled" do
      line = importer.lines.first
      expect(line[:first_name]).to eq("James")
    end
  end

  describe "parser" do
    it "should parse the last_name" do
      line = importer.lines.first
      expect(line[:last_name]).to eq("BOND")
    end
  end

  describe "required headers" do

    it "should include header if its a regex" do
      expect(importer.required_headers).to include(/^Email$/i)
    end

    it "should include header if its a string" do
      expect(importer.required_headers).to include(/Prénom/i)
    end

    it "should not require if header is a hash with a key required set to false" do
      expect(importer.required_headers).to_not include(/Actif/i)
    end
  end

  describe "-transform_header_to_regexp" do
    it "should transform a string to a regexp" do
      expect(importer.send(:transform_header_to_regexp, "Email")).to eq(/^Email$/i)
    end

    it "should not transform a regexp" do
      expect(importer.send(:transform_header_to_regexp, /^Email$/i)).to eq(/^Email$/i)
    end
  end
end