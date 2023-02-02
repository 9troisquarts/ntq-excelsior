# NtqExcelsior

Ntq excelsior simplifie l'import et l'export de vos données.

## Installation

Install the gem and add to the application's Gemfile by executing:

    $ bundle add ntq_excelsior

If bundler is not being used to manage dependencies, install the gem by executing:

    $ gem install ntq_excelsior

## Usage

### Export

```ruby
class UserExporter < NtqExcelsior::Exporter

  schema ({
    name: 'Mobilités',
    extra_headers: [
      [
        {
          title: "Utilisateurs",
          width: 4
        }
      ],
    ],
    columns: [
      {
        title: 'Name',
        resolve: -> (record) { [record.first_name, record.last_name].join(' ') },
      },
      {
        title: 'Email',
        resolve: 'email'
      },
      {
        title: 'Address (nested)',
        resolve: ['address', 'address_one']
      },
      {
        title: 'City (nested)',
        resolve: ['address', 'city']
      }
    ]
  })

end

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
