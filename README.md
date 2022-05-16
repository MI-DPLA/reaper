# reaper
Ruby program to convert microsoft excel files into Dublin Core XML for import into Combine collections.

Reaper converts each tab in an excel spreadsheet of DC data into an dc:xml file.

## Set up
- Use Ruby 2.7 or above
- Reaper uses the 'xsv' gem to read the spreadsheet data. Use the 'gem install xsv' command in your reaper directory.
- It also uses the ruby 'optparse' library to process the command line options you enter, this is included as part of Ruby.
## Data Prep
- Use the included dc_data.xlsx to prepare your tabular data for conversion to Dublin Core xml. 
- Your data can be in multiple tabs in Excel and each tab will be converted to a new xml file when you run reaper.rb with ruby. 
- Keep the dc_data.xlsx file in the same directory as the reaper.rb program.
- You can the dc_data.xlsx file and rename it, just use the new name as the input file for reaper.
## Running Reaper
- Use 'ruby reaper.rb -f dc_data.xlsx' to process the data. The excel tab names will be the resulting xml file names.
- Import the xml file directly into Combine as static xml.
