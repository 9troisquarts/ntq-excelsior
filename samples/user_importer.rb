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
    active: {
      header: /^Actif.*/i,
      humanized_header: 'Actif', # (Optional) if provided, will be displayed instead regex in missing headers
      required: false
    },
    last_name: {
      header: /^Nom$/i,
      required: true,
      parser: ->(value) { value&.upcase }
    },

  })

  def import_line(line, save: true)
    super do |record, line|
      record.email = line[:email]
      record.first_name = line[:first_name]
      record.last_name = line[:last_name]
    end
  end
end
