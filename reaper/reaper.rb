require 'xsv'
require 'optparse'

# print("Please place the Dublin Core Spreadsheet 'dc_data.xlsx' file in the same directory as the program ruby program reaper.rb.")
# puts
# puts

def reaper_run(spreadsh_data_file)

  intro_text = '<?xml version="1.0" encoding="UTF-8"?><documents>'
  extro_text = '</documents>'

  row_intro = '<oai_dc:dc xmlns:dcterms="http://purl.org/dc/dc/terms/" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:oai_dc="http://www.openarchives.org/OAI/2.0/oai_dc/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://www.openarchives.org/OAI/2.0/" xsi:schemaLocation="http://www.openarchives.org/OAI/2.0/oai_dc/ http://www.openarchives.org/OAI/2.0/oai_dc.xsd">'
  row_extro = '</oai_dc:dc>'
  open_tags = ["<dc:title>", "<dc:rights>", "<dc:publisher>", "<dc:identifier>", "<dc:identifier>", "<dc:identifier>", "<dcterms:isPartOf>", "<dc:language>", "<dc:type>", "<dc:date>", "<dc:coverage>", "<dc:subject>", "<dc:creator>", "<dc:format>", "<dc:publisher>", "<dc:contributor>", "<dc:format>", "<dc:identifier>", "<dcterms:description>"]
  close_tags = ["</dc:title>", "</dc:rights>", "</dc:publisher>", "</dc:identifier>", "</dc:identifier>", "</dc:identifier>", "</dcterms:isPartOf>", "</dc:language>", "</dc:type>", "</dc:date>", "</dc:coverage>", "</dc:subject>", "</dc:creator>", "</dc:format>", "</dc:publisher>", "</dc:contributor>", "</dc:format>", "</dc:identifier>", "</dcterms:description>"]

  workbook_data = Xsv::Workbook.open(spreadsh_data_file)

  # Use only the first part of the spreadsheet filename for the resulting xml files
  workbook_file_name_array = []
  workbook_file_name_array = spreadsh_data_file.split(".")
  workbook_file_name_part = workbook_file_name_array[0]

  # Create a base filename prefix used for each xml output file
  xml_base_file_name = workbook_file_name_part.to_s

  # Get sheet names and put them into output xml file names
  xml_output_file_names = []
  sheet_names = workbook_data.sheets.map(&:name)

  sheet_names.each_with_index do |val, index|
    val = "sheet_number_" + index.to_s if ( val.nil? || val.strip.empty? )
    xml_output_file_names[index] = xml_base_file_name + "_" + val + ".xml"
  end
  
  # Step through all the worksheets in the file
  workbook_data.sheets.each_with_index do |sheet, sheet_num|

    File.open(xml_output_file_names[sheet_num], 'w') do |file|

    # write the opening text for the xml file
    file.puts intro_text

    # Iterate over rows in worksheet, skiping the column headers (rows 1 to 3)
    row_number = 0

    sheet.each_row do |row|
      row_number += 1
      next if row_number < 4
        file.puts row_intro
        row.each_with_index do |val, index| 
          file.puts "  #{open_tags[index]}#{val}#{close_tags[index]}" if ( !val.nil? && ( val.to_s.count("a-zA-Z0-9") > 0 ) )
        end
        file.puts row_extro
        # Add a line after each row's xml
        file.puts
      end

      # write the closing text for the xml file
      file.puts extro_text
    end
  end
end


# reaper_run('dc_data.xlsx')

# This will hold the options we parse
options = {}

OptionParser.new do |parser|
  parser.banner = "Usage: reaper.rb [options]"

  parser.on("-h", "--help", "Show this help message") do ||
      puts parser
  end

  # Whenever we see -n or --name, with an
  # argument, save the argument.
  parser.on("-f", "--file FILENAME.xlsx", "The file with the spreadsheet tabs to process into xml.") do |v|
    options[:file] = v
  end

end.parse!

# Now we can use the options hash however we like.

if options[:file].to_s == ''
  puts "You must include a excel data file"
  puts "Use 'ruby reaper.rb -f <excel input file>' to process data" 
elsif options[:file]
  puts "Processing data from file #{options[:file]}."
  reaper_run(options[:file])
end
