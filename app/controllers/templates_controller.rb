require 'rubygems'
require 'roo'

class TemplatesController < ApplicationController

	# GET /generate_template
	# GET /generate_template.json
	def index
 
		# Load the source Excel file, with all the special products info
		source = Excelx.new("http://beta.topart.com/csv/special.xlsx")
		source.default_sheet = source.sheets.first
		
		# Load the Magento template, which is in Open Office format
		template = Openoffice.new("http://beta.topart.com/csv/template.ods")
		template.default_sheet = template.sheets.first
		
		# Fill every line in the template file up with
		# the right value taken from the source input file
		2.upto(989) do |line|
		
			template.set(line, 'BV', "#{source.cell(line,'A')}")
			template.set(line, 'A', "#{source.cell(line,'B')}")
			template.set(line, 'AR', "#{source.cell(line,'B')}")
			template.set(line, 'CM', "#{source.cell(line,'C')}")
			template.set(line, 'CK', "#{source.cell(line,'D')}")
			template.set(line, 'AT', "#{source.cell(line,'E')}")
			template.set(line, 'CL', "#{source.cell(line,'F')}")
			template.set(line, 'AU', "#{source.cell(line,'G')}")
			template.set(line, 'CI', "#{source.cell(line,'H')}")
			template.set(line, 'L', "#{source.cell(line,'I')}")
			template.set(line, 'M', "#{source.cell(line,'J')}")
			# we got until "artist_last_name"
			
			template.set(line, 'BG', "#{source.cell(line,'K')}")
			template.set(line, 'CQ', "#{source.cell(line,'K')}")
			template.set(line, 'AB', "#{source.cell(line,'K')}")
			template.set(line, 'CC', "#{source.cell(line,'K')}")
			
			template.set(line, 'BU', "#{source.cell(line,'L')}")
			template.set(line, 'BP', "#{source.cell(line,'M')}")
			template.set(line, 'BQ', "#{source.cell(line,'N')}")
			template.set(line, 'AP', "#{source.cell(line,'O')}")
			template.set(line, 'AQ', "#{source.cell(line,'P')}")
			template.set(line, 'AN', "#{source.cell(line,'Q')}")
			template.set(line, 'CB', "#{source.cell(line,'Q')}")
			template.set(line, 'H', "#{source.cell(line,'S')}")
			# we got until "alt_size_1"
			
			template.set(line, 'I', "#{source.cell(line,'T')}")
			template.set(line, 'J', "#{source.cell(line,'U')}")
			template.set(line, 'K', "#{source.cell(line,'V')}")
			template.set(line, 'BM', "#{source.cell(line,'W')}")
			template.set(line, 'BY', "#{source.cell(line,'X')}")
			template.set(line, 'AE', "#{source.cell(line,'Y')}")
			template.set(line, 'CD', "#{source.cell(line,'Z')}")
			template.set(line, 'AH', "#{source.cell(line,'AA')}")
			template.set(line, 'AX', "#{source.cell(line,'AB')}")
			template.set(line, 'AM', "#{source.cell(line,'AC')}")
			# we got until "ignore_discount"
			
			template.set(line, 'T', "#{source.cell(line,'AD')}")
			template.set(line, 'BN', "#{source.cell(line,'AE')}")
			template.set(line, 'AI', "#{source.cell(line,'AF')}")
			template.set(line, 'CA', "#{source.cell(line,'AG')}")
			template.set(line, 'O', "#{source.cell(line,'AH')}")
			template.set(line, 'BH', "#{source.cell(line,'AI')}")
			template.set(line, 'BX', "#{source.cell(line,'AJ')}")
			template.set(line, 'AS', "#{source.cell(line,'AK')}")
			template.set(line, 'Y', "#{source.cell(line,'AL')}")
			template.set(line, 'S', "#{source.cell(line,'AM')}")
			# we got until "created_by"
	  
			template.set(line, 'AA', "#{source.cell(line,'AN')}")
			template.set(line, 'BC', "#{source.cell(line,'AO')}")
			template.set(line, 'Z', "#{source.cell(line,'AP')}")
			template.set(line, 'AG', "#{source.cell(line,'AQ')}")
			template.set(line, 'AC', "#{source.cell(line,'AR')}")
			template.set(line, 'BR', "#{source.cell(line,'AS')}")
			template.set(line, 'BT', "#{source.cell(line,'AT')}")
			template.set(line, 'BS', "#{source.cell(line,'AU')}")
			template.set(line, 'BK', "#{source.cell(line,'AV')}")
			template.set(line, 'CF', "#{source.cell(line,'AW')}")
			template.set(line, 'AD', "#{source.cell(line,'AX')}")
			# we got until "do_not_display"
			
			template.set(line, 'C', "Topart - Special Products")
			template.set(line, 'D', "simple")
			template.set(line, 'E', "Collections/Oscar Night")
			template.set(line, 'F', "Root Category")
			template.set(line, 'G', "base")
			template.set(line, 'AL', "0")
			template.set(line, 'BE', "Use config")
			template.set(line, 'BF', "Use config")
			template.set(line, 'BL', "Block after Info Column")
			template.set(line, 'CJ', "1")
			
			template.set(line, 'BZ', "1")
			template.set(line, 'CV', "1")
			template.set(line, 'CU', "4")
			template.set(line, 'CN', "0")
	  
		end
		
		template.to_csv("filled_template.csv")
 
		respond_to do |format|
			format.html # index.html.erb
		end

	end

end
