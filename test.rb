require "google_drive"

class GoogleSheet
attr_reader :worksheet

    def initialize(session, worksheet_title)
    
        @worksheet = session.spreadsheet_by_title(worksheet_title).worksheets[0]
    end
        
    def table_values
        (1..@worksheet.num_rows).map do |row|
          (1..@worksheet.num_cols).map do |col|
            @worksheet[row, col]
          end
        end
      end
    
      def row_by_index(index)
        (1..@worksheet.num_cols).map do |col|
          @worksheet[index, col]
        end
      end

      include Enumerable

  def each(&block)
    (1..@worksheet.num_rows).each do |row|
      (1..@worksheet.num_cols).each do |col|
        block.call(@worksheet[row, col])
      end
    end
  end

  def merged_cell?(row, col)
    @worksheet.merged_ranges.any? do |merged_range|
      # Check if the given cell (row, col) is within any merged range
      merged_range.include?(row, col)
    end
  end

  

  
end


class Table
  def initialize(worksheet)
    @worksheet = worksheet
  end

  #Biblioteka prepoznaje ukoliko postoji na bilo koji način ključna reč total ili subtotal 
  #unutar sheet-a, i ignoriše taj red

  #The column_values method has been modified to filter out rows containing the keywords 'total' or 'subtotal' in the specified column.
  #It uses reject to remove rows that include either 'total' or 'subtotal' (case insensitive) in the column.
  #The map method is then used to retrieve the values from the remaining rows.
  #compact is used to remove any nil values that might occur after mapping.
  
  def column_values(column_name)
    header_row = @worksheet.rows.first
    col_index = header_row.index(column_name)
      
    if col_index
      column = @worksheet.rows.drop(1).reject do |row|
        row[col_index]&.downcase&.include?('total') || row[col_index]&.downcase&.include?('subtotal')
      end.map { |row| row[col_index] }
      column.compact
    else
      nil
    end
  end

  # Access value in a specific column at a given index
  def value_at(column_name, index)
    header_row = @worksheet.rows.first
    col_index = header_row.index(column_name)

    if col_index
      row = @worksheet.rows[index + 1] # Skip the header row
      row[col_index] if row
    else
      nil
    end
  end

  def set_value_at(column_name, index, value)
    header_row = @worksheet.rows.first
    col_index = header_row.index(column_name)
  
    if col_index
      cell = @worksheet[index + 1, col_index + 1]  # Adjust index by 1 to match worksheet indexing
  
      # Check if the cell exists and isn't empty before assigning the value
      if cell && !cell.empty?
        cell.value = value
        @worksheet.save # Save changes to the worksheet
      else
        puts "Cell is empty or doesn't exist." # Handle the case where the cell is empty or doesn't exist
      end
    else
      puts "Column '#{column_name}' not found." # Handle the case where the column doesn't exist
    end
  end


  # Access entire column directly using method names like prvaKolona, drugaKolona, etc.
  def prvaKolona
    column_values('Prva Kolona')
  end

  def drugaKolona
    column_values('Druga Kolona')
  end

  def trecaKolona
    column_values('Treca Kolona')
  end

  def sum(column_name)
    column = column_values(column_name)
    
    # Filter out non-numeric values before calculating the sum
    numeric_values = column.compact.map(&:to_i).select { |value| value.to_s == column_name.to_s }
  
    numeric_values.sum
  end

  def avg(column_name)
    column = column_values(column_name)
    
    # Filter out non-numeric values before calculating the average
    numeric_values = column.compact.map(&:to_f).select { |value| value.to_s == column_name.to_s }
  
    if numeric_values.any?
      sum = numeric_values.sum
      count = numeric_values.size
      sum / count
    else
      nil # Return nil if there are no numeric values
    end
  end

  # Retrieve individual row based on a cell value in a specific column
  def row_by_cell_value(column_name, value)
    header_row = @worksheet.rows.first
    col_index = header_row.index(column_name)

    if col_index
      row = @worksheet.rows.find { |r| r[col_index] == value }
      row
    else
      nil
    end
  end

  def map_column(column_name, &block)
    column = column_values(column_name)

    # Filter out non-numeric values before mapping
    numeric_values = column.compact.map(&:to_i).select { |value| value.to_s == column_name.to_s }

    numeric_values.map(&block) if numeric_values.any?
  end

  def select_column(column_name, &block)
    column = column_values(column_name)
    column.select(&block) if column
  end

  def reduce_column(column_name, initial, &block)
    column = column_values(column_name)
    column.compact.reduce(initial, &block) if column
  end

end

def main
 session = GoogleDrive::Session.from_config("config.json")
 ws = session.spreadsheet_by_key("11-bdcDOK2UmKPYIJLjytbxTyW6WTJekPlfF8O0VtSjM").worksheets[0]
 gs = GoogleSheet.new(session, 'Rubyproject')
 table = Table.new(gs.worksheet)

 puts "Entire column 'Prva kolona':"
  p table.column_values('Prva kolona') # Access entire column

  puts "Accessing value at index 1 in 'Prva kolona':"
  p table.value_at('Prva kolona', 1) # Access value in column at index 1


  puts 'Table values:'
  p gs.table_values

  puts 'Row by index (e.g., 3rd row):'
  p gs.row_by_index(3)

  puts 'All cells in the worksheet:'
  gs.each { |cell| puts cell }

# Access entire columns directly using method names
prva_kolona_values = table.prvaKolona
druga_kolona_values = table.drugaKolona
treca_kolona_values = table.trecaKolona

puts "Values in 'Prva kolona': #{prva_kolona_values}"
puts "Values in 'Druga kolona': #{druga_kolona_values}"
puts "Values in 'Treca kolona': #{treca_kolona_values}"

# Accessing specific values within columns
value_in_druga_kolona = table.value_at('Druga kolona', 0) # Accessing first value in "Druga kolona"
p value_in_druga_kolona # Output: 25

# Setting a new value at a specific index in a column
table.set_value_at('Prva kolona', 1, 8) # Setting a new value '8' at index 1 in "Prva kolona"

updated_prva_kolona = table.prvaKolona
p updated_prva_kolona # Output: [1, 8] - Updated values from "Prva kolona"


sum_prva_kolona = table.sum('Prva kolona')
avg_druga_kolona = table.avg('Druga kolona')

puts "Sum of 'Prva kolona': #{sum_prva_kolona}"
puts "Average of 'Druga kolona': #{avg_druga_kolona}"

specific_row = table.row_by_cell_value('Prva kolona', '1')
puts "Row with 'Prva kolona' value '1': #{specific_row}"

# Using map, select, reduce on columns
mapped_treca_kolona = table.map_column('Treca kolona') { |value| value * 2 }
selected_prva_kolona = table.select_column('Prva kolona') { |value| value.to_i > 2 }
reduced_druga_kolona = table.reduce_column('Druga kolona', 0) { |sum, value| sum + value.to_i }

puts "Mapped values in 'Treca kolona': #{mapped_treca_kolona}"
puts "Selected values in 'Prva kolona' greater than 2: #{selected_prva_kolona}"
puts "Reduced value in 'Druga kolona': #{reduced_druga_kolona}"


end

main if _FILE_ == $PROGRAM_NAME