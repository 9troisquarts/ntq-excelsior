module NtqExcelsior
  module Exporters
    # Handles the generation and management of Excel headers.
    # This class is responsible for creating and formatting header rows in Excel exports.
    # It supports nested headers, column merging, and style application.
    #
    # == Example
    #   columns = [
    #     Column.new(title: "Name", styles: [:bold]),
    #     Column.new(title: "Details", children: [
    #       Column.new(title: "Age"),
    #       Column.new(title: "City")
    #     ])
    #   ]
    #   header = Header.new(columns, worksheet_styles: worksheet_styles)
    #   rows = header.resolve_rows
    #
    class Header
      include CellHelper

      # Structure d'une ligne vide utilisée comme base pour toutes les lignes
      # Contient les tableaux et propriétés nécessaires pour une ligne Excel
      EMPTY_ROW = {
        values: [],           # Valeurs des cellules
        styles: [],           # Styles appliqués aux cellules
        merge_cells: [],      # Cellules à fusionner
        data_validations: [], # Validations de données
        height: nil           # Hauteur de la ligne
      }.freeze

      # Attributs de la classe
      # @columns: Définition des colonnes à générer
      # @worksheet_styles: Styles disponibles pour le worksheet
      # @offset_row: Décalage en nombre de lignes pour le placement des headers
      attr_reader :columns, :worksheet_styles, :offset_row

      # Initialise une nouvelle instance de Header
      # @param columns [Array<Column>] Les colonnes à partir desquelles générer les headers
      # @param offset_row [Integer] Décalage en nombre de lignes (défaut: 1)
      # @param worksheet_styles [WorksheetStyles] Styles disponibles pour le worksheet
      def initialize(columns, offset_row = 1, worksheet_styles: nil)
        raise ArgumentError, "columns must be an Array" unless columns.is_a?(Array)
        raise ArgumentError, "offset_row must be a positive integer" unless offset_row.is_a?(Integer) && offset_row >= 0

        @columns = columns
        @offset_row = offset_row
        @worksheet_styles = worksheet_styles
      end

      # Résout la configuration des lignes d'en-tête
      # Génère la structure complète des headers en tenant compte de la profondeur et du contexte
      # @param context [Object] Contexte optionnel pour la visibilité des colonnes
      # @return [Array<Hash>] Configuration des lignes d'en-tête
      def resolve_rows(context: nil)
        return [empty_row] if columns.empty?

        max_depth = columns.map(&:depth).max || 1
        initial_rows = Array.new(max_depth) { empty_row }

        normalize_rows(place_headers(initial_rows, max_depth, context: context))
      end

      private

      # Méthodes de base pour la création de structures

      # Crée une nouvelle ligne vide basée sur EMPTY_ROW
      # @return [Hash] Nouvelle ligne vide
      def empty_row
        EMPTY_ROW.dup
      end

      # Duplique un ensemble de lignes en créant de nouvelles instances
      # @param rows [Array<Hash>] Lignes à dupliquer
      # @return [Array<Hash>] Nouvelles instances des lignes
      def duplicate_rows(rows)
        rows.map do |r|
          next empty_row unless r

          {
            values: r[:values]&.dup || [],
            styles: r[:styles]&.dup || [],
            merge_cells: r[:merge_cells]&.dup || [],
            data_validations: r[:data_validations]&.dup || [],
            height: r[:height]
          }
        end
      end

      # Méthodes de placement des headers

      # Place tous les headers dans la structure de lignes
      # @param rows [Array<Hash>] Configuration initiale des lignes
      # @param depth [Integer] Profondeur maximale des headers imbriqués
      # @param context [Object] Contexte pour la visibilité des colonnes
      # @return [Array<Hash>] Lignes avec les headers placés
      def place_headers(rows, depth, context: nil)
        columns.reduce({ next_col: 0, rows: rows }) do |acc, column|
          result = place_header(column, acc[:next_col], 0, acc[:rows], depth, context: context)
          { next_col: result[:next_col], rows: result[:rows] }
        end[:rows]
      end

      # Place un header individuel dans la structure
      # Gère la visibilité et le type de header (parent ou feuille)
      # @param column [Column] Colonne à placer
      # @param start_col [Integer] Index de colonne de départ
      # @param row [Integer] Index de ligne
      # @param rows [Array<Hash>] Configuration actuelle des lignes
      # @param depth [Integer] Profondeur maximale
      # @param context [Object] Contexte pour la visibilité
      # @return [Hash] Prochaine colonne et lignes mises à jour
      def place_header(column, start_col, row, rows, depth, context: nil)
        return { next_col: start_col, rows: rows } unless column.visible?(context: context)

        if column.has_children?
          place_parent_header(column, start_col, row, rows, depth, context: context)
        else
          place_leaf_header(column, start_col, row, rows, depth)
        end
      end

      # Place un header parent avec ses enfants
      # Gère la fusion des cellules pour les headers parents
      # @param column [Column] Colonne parent
      # @param start_col [Integer] Index de colonne de départ
      # @param row [Integer] Index de ligne
      # @param rows [Array<Hash>] Configuration actuelle des lignes
      # @param depth [Integer] Profondeur maximale
      # @param context [Object] Contexte pour la visibilité
      # @return [Hash] Prochaine colonne et lignes mises à jour
      def place_parent_header(column, start_col, row, rows, depth, context: nil)
        # Place d'abord les enfants
        result = column.children.reduce({ next_col: start_col, rows: rows }) do |acc, child|
          result = place_header(child, acc[:next_col], row + 1, acc[:rows], depth, context: context)
          { next_col: result[:next_col], rows: result[:rows] }
        end

        child_start_col = result[:next_col]
        rows_with_children = result[:rows]
        current_width = child_start_col - start_col

        # Place le header parent et fusionne les cellules si nécessaire
        rows_with_cell = place_cell(column, start_col, row, rows_with_children)
        is_last = row == depth
        final_rows = if current_width > 1 && column.children.any?
                       merge_cells(rows_with_cell, row, start_col, current_width, :parent)
                     else
                       rows_with_cell
                     end
        { next_col: child_start_col, rows: final_rows }
      end

      # Place un header feuille (sans enfants)
      # Gère la fusion verticale des cellules si nécessaire
      # @param column [Column] Colonne feuille
      # @param start_col [Integer] Index de colonne de départ
      # @param row [Integer] Index de ligne
      # @param rows [Array<Hash>] Configuration actuelle des lignes
      # @param depth [Integer] Profondeur maximale
      # @return [Hash] Prochaine colonne et lignes mises à jour
      def place_leaf_header(column, start_col, row, rows, depth)
        rows_with_cell = place_cell(column, start_col, row, rows, validation: true)
        final_rows = if row < depth - 1 && column.children.any?
                       merge_cells(rows_with_cell, row, start_col, depth - row,
                                   :leaf)
                     else
                       rows_with_cell
                     end
        { next_col: start_col + 1, rows: final_rows }
      end

      # Méthodes de manipulation des cellules

      # Place une cellule d'en-tête avec sa valeur et ses styles
      # @param column [Column] Colonne contenant les informations d'en-tête
      # @param col [Integer] Index de colonne
      # @param row [Integer] Index de ligne
      # @param rows [Array<Hash>] Configuration actuelle des lignes
      # @param validation [Boolean] Si true, ajoute les validations de données
      # @return [Array<Hash>] Lignes mises à jour avec la nouvelle cellule
      def place_cell(column, col, row, rows, validation: false)
        new_rows = rows.dup
        new_rows[row] ||= empty_row
        new_row = ensure_cell_space(new_rows[row], col)

        # Place la valeur et les styles
        new_row[:values][col] = column.title || ""
        new_row[:styles][col] = get_styles(column.header_styles)

        # Ajoute les validations si nécessaire
        if validation && column.list
          range = cells_range([col + 1, row + 1], [col + 1, 1_000_000])
          new_row[:data_validations] << {
            range: range,
            config: column.validations
          }
        end

        new_rows[row] = new_row
        new_rows
      end

      # Fusionne des cellules pour un header
      # @param rows [Array<Hash>] Configuration actuelle des lignes
      # @param row [Integer] Index de ligne
      # @param start_col [Integer] Index de colonne de départ
      # @param span [Integer] Nombre de cellules à fusionner
      # @param type [Symbol] Type de fusion (:parent pour horizontal, :leaf pour vertical)
      # @return [Array<Hash>] Lignes mises à jour avec les fusions
      def merge_cells(rows, row, start_col, span, type)
        return rows if span <= 1

        new_rows = rows
        new_rows[row] = (rows[row] || empty_row).dup
        new_rows[row][:merge_cells] = (new_rows[row][:merge_cells] || []).dup

        # Détermine la plage de fusion selon le type
        merge_range = case type
                      when :parent
                        cells_range([start_col + 1, offset_row + row], [start_col + span, offset_row + row])
                      when :leaf
                        cells_range([start_col + 1, offset_row + row], [start_col + 1, offset_row + row + span - 1])
                      end

        new_rows[row][:merge_cells] << merge_range
        clear_merged_cells(new_rows, row, start_col, span, type)
      end

      # Efface le contenu et les styles des cellules qui seront fusionnées
      # @param rows [Array<Hash>] Configuration actuelle des lignes
      # @param row [Integer] Index de ligne de départ
      # @param start_col [Integer] Index de colonne de départ
      # @param span [Integer] Nombre de cellules à effacer
      # @param type [Symbol] Type de fusion (:parent ou :leaf)
      # @return [Array<Hash>] Lignes mises à jour avec les cellules effacées
      def clear_merged_cells(rows, row, start_col, span, type)
        new_rows = duplicate_rows(rows)

        case type
        when :parent
          # Efface les cellules horizontalement
          new_row = new_rows[row] || empty_row
          new_values = new_row[:values]&.dup || []
          new_styles = new_row[:styles]&.dup || []

          (1...span).each do |i|
            new_values[start_col + i] = nil
            new_styles[start_col + i] = {}
          end

          new_rows[row] = new_row.merge(values: new_values, styles: new_styles)
        when :leaf
          # Efface les cellules verticalement
          ((row + 1)...(row + span)).each do |r|
            new_rows[r] = empty_row if new_rows[r].nil?
            new_row = new_rows[r]
            new_values = new_row[:values]&.dup || []
            new_styles = new_row[:styles]&.dup || []

            new_values[start_col] = nil
            new_styles[start_col] = {}

            new_rows[r] = new_row.merge(values: new_values, styles: new_styles)
          end
        end

        new_rows
      end

      # Méthodes utilitaires

      # Normalise la longueur des lignes pour assurer un nombre égal de colonnes
      # @param rows [Array<Hash>] Lignes à normaliser
      # @return [Array<Hash>] Lignes normalisées
      def normalize_rows(rows)
        max_length = rows.map { |row| row[:values].length }.max

        rows.map! do |row|
          row[:values] = pad_array(row[:values], max_length, nil)
          row[:styles] = pad_array(row[:styles], max_length, {})
          row
        end
      end

      # Complète un tableau jusqu'à une longueur spécifiée
      # @param array [Array] Tableau à compléter
      # @param length [Integer] Longueur souhaitée
      # @param value [Object] Valeur de remplissage
      # @return [Array] Tableau complété
      def pad_array(array, length, value)
        return array if array.length >= length

        array.fill(value, array.length...length)
        array
      end

      # Assure qu'une ligne a assez d'espace pour une cellule
      # @param row [Hash] Ligne à modifier
      # @param col [Integer] Index de colonne
      # @return [Hash] Ligne mise à jour
      def ensure_cell_space(row, col)
        base_row = row || empty_row

        base_row[:values] = ensure_size(base_row[:values], col + 1, nil)
        base_row[:styles] = ensure_size(base_row[:styles], col + 1, {})

        base_row
      end

      # Assure qu'une ligne a assez d'espace pour une fusion
      # @param row [Hash] Ligne à modifier
      # @param start_col [Integer] Index de colonne de départ
      # @param span [Integer, Range] Étendue de la fusion
      # @return [Hash] Ligne mise à jour
      def ensure_merge_space(row, start_col, span)
        size = case span
               when Integer then start_col + span
               when Range then span.end
               end

        ensure_cell_space(row, size)
      end

      # Assure qu'un tableau a au moins une taille spécifiée
      # @param array [Array] Tableau à modifier
      # @param size [Integer] Taille minimale requise
      # @param default_value [Object] Valeur par défaut
      # @return [Array] Tableau mis à jour
      def ensure_size(array, size, default_value)
        return array if array.size >= size

        array + Array.new(size - array.size, default_value)
      end

      # Assure qu'un tableau de lignes a au moins une taille spécifiée
      # @param rows [Array<Hash>] Tableau de lignes
      # @param min_size [Integer] Taille minimale requise
      # @return [Array<Hash>] Tableau de lignes mis à jour
      def ensure_rows(rows, min_size)
        return rows if rows.size >= min_size

        rows + Array.new(min_size - rows.size) { empty_row }
      end

      # Obtient les styles combinés pour un header
      # Utilise un cache pour éviter de recalculer les styles
      # @param styles [Array<Symbol>] Clés de style à appliquer
      # @return [Hash] Styles combinés
      def get_styles(styles)
        return {} if styles.nil? || styles.empty? || !worksheet_styles

        @style_cache ||= {}
        cache_key = styles.hash

        @style_cache[cache_key] ||= styles.reduce({}) do |styles_hash, style_key|
          styles_hash.merge(worksheet_styles.get_style(style_key))
        end
      end
    end
  end
end
