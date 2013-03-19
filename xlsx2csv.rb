require 'roo'
require 'csv'


class Xlsx2Csv
  def initialize
    @infilename = INPUT
    @input = Roo::Spreadsheet.open(@infilename)
    @outfilename = File.basename(INPUT,".*")
    @output = CSV.open(OUTPUT+"/"+@outfilename+".csv","w")
    @sheet = SHEET
  end

  def xlsx_to_csv
    (@input.first_row..@input.last_row).each{|row_ind|
      @output_row = []
      (@input.first_column..@input.last_column).each{|column_ind|
        @output_row << @input.sheet(@sheet).cell(row_ind,column_ind)
      }
      @output << @output_row
    }
    @output.close
  end

  def valid?
    if @infilename.nil? || @infilename !~ /.xlsx$/
      raise "Input file must be XLSX file"
    end
  end

  def main
    if valid?
      xlsx_to_csv
    end
  end
end


if __FILE__ == $0
  if ARGV.length != 2 && ARGV.length != 3
    puts "Usage: ruby xlsx2csv INPUT_FILENAME OUTPUT_DIRNAME SHEET_NUMBER(optional)"
    return
  end
  INPUT = ARGV[0]
  OUTPUT = ARGV[1]
  if ARGV.length == 3 && !ARGV[2].nil?
    SHEET = ARGV[2]
  else
    SHEET = 1
  end
  Xlsx2Csv.new.main
end
