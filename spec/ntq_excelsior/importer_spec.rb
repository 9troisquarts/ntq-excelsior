require "spec_helper"
require_relative "../../samples/user_importer"

RSpec.describe NtqExcelsior::Importer do
  let(:file_path) { File.expand_path("../fixtures/users.xlsx", __dir__) }

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

    it "should ignore additional columns not present in schema" do
      line = importer.lines.first
      expect(line).to_not have_key(:telephone)
      expect(line).to_not have_key(:adresse)
    end
  end

  describe "fixture headers" do
    it "should include the new Telephone and Adresse headers with optional markers" do
      headers = importer.spreadsheet.sheet(importer.spreadsheet.sheets[0]).row(1)
      expect(headers).to include("Telephone *")
      expect(headers).to include("Adresse*")
    end
  end

  describe "required headers" do
    it "should include header if its a regex from the schema" do
      expect(importer.required_headers).to include(/Prénom/i)
    end

    it "should build a regexp for string headers that allows an optional * suffix" do
      email_regexp = importer.required_headers.find { |h| h.is_a?(Regexp) && h.match?("Email") }
      expect(email_regexp).to be_a(Regexp)
      expect("Email").to match(email_regexp)
      expect("Email *").to match(email_regexp)
    end

    it "should not require if header is a hash with a key required set to false" do
      expect(importer.required_headers).to_not include(/Actif/i)
    end
  end

  describe "-transform_header_to_regexp" do
    it "should transform a string to a regexp" do
      transformed = importer.send(:transform_header_to_regexp, "Email")
      expect(transformed).to be_a(Regexp)
      expect("Email").to match(transformed)
      expect("Email *").to match(transformed)
    end

    it "should not transform a regexp" do
      expect(importer.send(:transform_header_to_regexp, /^Email$/i)).to eq(/^Email$/i)
    end
  end

  describe "normalize_required_marker" do
    it "strips trailing optional asterisk marker" do
      expect(importer.send(:normalize_required_marker, "Nom *")).to eq("Nom")
    end

    it "strips trailing optional asterisk marker without spaces" do
      expect(importer.send(:normalize_required_marker, "Adresse*")).to eq("Adresse")
    end
  end

  describe "humanize_missing_header" do
    it "converts missing header regexp to readable header label" do
      expect(importer.send(:humanize_missing_header, "/^Keywords(?:\\s*\\*)?$/i")).to eq("Keywords")
    end
  end

  describe "spreadsheet_data error messages" do
    let(:sheet_double) { double("sheet") }
    let(:spreadsheet_double) { double("spreadsheet") }

    before do
      allow(importer).to receive(:spreadsheet).and_return(spreadsheet_double)
      allow(spreadsheet_double).to receive(:sheets).and_return(["Sheet1"])
      allow(spreadsheet_double).to receive(:sheet).with("Sheet1").and_return(sheet_double)
    end

    it "humanizes missing string headers from Roo message" do
      allow(sheet_double).to receive(:parse).and_raise(
        Roo::HeaderRowNotFoundError,
        "The following headers are missing:: /^Email(?:\\s*\\*)?$/i"
      )

      expect { importer.spreadsheet_data }.to raise_error(Roo::HeaderRowNotFoundError, "Email")
    end

    it "humanizes missing regex headers when no humanized_header is configured" do
      allow(sheet_double).to receive(:parse).and_raise(
        Roo::HeaderRowNotFoundError,
        "The following headers are missing:: /^Nom(?:\\s*\\*)?$/i"
      )

      expect { importer.spreadsheet_data }.to raise_error(Roo::HeaderRowNotFoundError, "Nom")
    end
  end
end
