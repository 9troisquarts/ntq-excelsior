# frozen_string_literal: true

require_relative "lib/ntq_excelsior/version"

Gem::Specification.new do |spec|
  spec.name = "ntq_excelsior"
  spec.version = NtqExcelsior::VERSION
  spec.authors = ["Kevin"]
  spec.email = ["kevin@9troisquarts.com"]

  spec.summary = "Library use by 9tq for import/export"
  spec.description = "Library use by 9tq for import/export"
  spec.homepage = "https://github.com/9troisquarts/ntq-excelsior"
  spec.license = "MIT"
  spec.required_ruby_version = ">= 2.6.0"

  spec.metadata["homepage_uri"] = spec.homepage
  spec.metadata["source_code_uri"] = "https://github.com/9troisquarts/ntq-excelsior"
  spec.metadata["changelog_uri"] = "https://github.com/9troisquarts/ntq-excelsior/CHANGELOG.md"

  # Specify which files should be added to the gem when it is released.
  # The `git ls-files -z` loads the files in the RubyGem that have been added into git.
  spec.files = Dir.chdir(__dir__) do
    `git ls-files -z`.split("\x0").reject do |f|
      (f == __FILE__) || f.match(%r{\A(?:(?:bin|test|spec|features)/|\.(?:git|travis|circleci)|appveyor)})
    end
  end
  spec.bindir = "exe"
  spec.executables = spec.files.grep(%r{\Aexe/}) { |f| File.basename(f) }
  spec.require_paths = ["lib"]

  # Uncomment to register a new dependency of your gem
  # spec.add_dependency "example-gem", "~> 1.0"
  spec.add_dependency "caxlsx", "< 4"
  spec.add_dependency "roo", "< 3"

  # For more information and examples about making a new gem, check out our
  # guide at: https://bundler.io/guides/creating_gem.html
end