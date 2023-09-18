# NtqExcelsior

[![Maintainability](https://api.codeclimate.com/v1/badges/8bc43b15a0a8bfc5d660/maintainability)](https://codeclimate.com/github/9troisquarts/ntq-excelsior/maintainability)

Ntq excelsior simplifie l'import et l'export de vos données.

## Installation

Install the gem and add to the application's Gemfile by executing:

    $ bundle add ntq_excelsior

If bundler is not being used to manage dependencies, install the gem by executing:

    $ gem install ntq_excelsior

## Usage

### Export

```ruby
# Exporter class
class UserExporter < NtqExcelsior::Exporter

  styles ({
    blue: {
      fg_color: "2F5496",
    }
  })

  schema ({
    name: 'Utilisateurs',
    extra_headers: [
      [
        {
          title: "Utilisateurs",
          width: -> (context) { context[:current_user].can?(:access_to_email, User) ? 4 : 3 }
          # width: 4,
          styles: [:bold]
        }
      ],
    ],
    columns: [
      {
        title: 'Name',
        resolve: -> (record) { [record.first_name, record.last_name].join(' ') },
        styles: [:bold, :blue]
      },
      {
        title: 'Email',
        header_styles: [:blue],
        resolve: 'email',
        visible: -> (record, context) { context[:current_user].can?(:access_to_email, User) }
      },
      {
        title: 'Birthdate',
        resolve: 'birthdate',
        visible: true # Optional
      }
      {
        title: 'Address (nested)',
        resolve: ['address', 'address_one']
      },
      {
        title: 'City (nested)',
        resolve: ['address', 'city']
      },
      {
        title: 'Age',
        resolve: 'age',
        type: :number
      }
    ]
  })

end

exporter = UserExporter.new(@users)
# Optional : Context can be passed to exporter
exporter.context = { current_user: current_user }
stream = exporter.export.to_stream.read

# In ruby file
File.open("export.xlsx", "w") do |tmp|
  tmp.binmode
  tmp.write(stream)
end

# In Controller action
send_data stream, type: 'application/xlsx', filename: "filename.xlsx"
```

### Multiple workbooks export

```ruby
  user_exporter = UserExporter.new(@user_data)
  product_exporter = ProductExporter.new(@product_data)

  exporter = NtqExcelsior::MultiWorkbookExporter.new([user_exporter, product_exporter])
  stream = exporter.export.to_stream.read

  # In ruby file
  File.open("export.xlsx", "w") do |tmp|
    tmp.binmode
    tmp.write(stream)
  end
  
  # In Controller action
  send_data stream, type: 'application/xlsx', filename: "filename.xlsx"
```

### Import

```ruby
class UserImporter < NtqExcelsior::Importer

  model_klass "User"

  primary_key :email

  ## if set to save, the record will not be save automatically after import_line
  ## Errors have to be added to the @errors hash manually
  # autosave false

  structure [{
    header: "Email",
    description: "Email de l'utilisateur a créer ou modifier",
    required: true,
  }, 
  {
    header: "Actif",
    description: "Utilisateur activé",
    required: true,
    values: [
      {
        header: "True",
        description: "Activé",
      }, {
        header: "False",
        description: "Inactivé",
      }
    ]
  }, {
    header: "Prénom",
    required: true,
  }, {
    header: "Nom",
    required: true,
  }]

  sample_file "/import_samples/users.xlsx"

  schema({
    email: 'Email',
    first_name: /Prénom/i,
    last_name: {
      header: /Nom/i,
      required: true
    },
    active: {
      header: /Actif.+/i,
      humanized_header: 'Actif', # (Optional) if provided, will be displayed instead regex in missing headers
      required: false
    }
  })

  def import_line(line, save: true)
    super do |record, line|
      record.email = line[:email]
      record.first_name = line[:first_name]
      record.last_name = line[:last_name]
    end
  end
end

importer = UserImporter.new
importer.file = file
result = importer.import
# { success_count: 2, error_lines: [] }
```

## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake spec` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and the created tag, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/[USERNAME]/ntq_excelsior. This project is intended to be a safe, welcoming space for collaboration, and contributors are expected to adhere to the [code of conduct](https://github.com/[USERNAME]/ntq_excelsior/blob/master/CODE_OF_CONDUCT.md).

## License

The gem is available as open source under the terms of the [MIT License](https://opensource.org/licenses/MIT).

## Code of Conduct

Everyone interacting in the NtqExcelsior project's codebases, issue trackers, chat rooms and mailing lists is expected to follow the [code of conduct](https://github.com/[USERNAME]/ntq_excelsior/blob/master/CODE_OF_CONDUCT.md).
