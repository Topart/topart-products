require 'rubygems'
require 'roo'

class TemplatesController < ApplicationController

	# GET /generate_template
	# GET /generate_template.json
	def index
 
		# Load the source Excel file, with all the special products info
		source = Excelx.new("http://beta.topart.com/csv/special_products.xlsx")
		source.default_sheet = source.sheets.first
		
		# Load the Magento template, which is in Open Office format
		template = Openoffice.new("http://beta.topart.com/csv/template.ods")
		template.default_sheet = template.sheets.first
		
		# Fill every line in the template file up with
		# the right value taken from the source input file
		
		#2.upto(989) do |line|
		@line = 2
		while @line < source.last_row
		
			template.set(@line, 'A', "#{source.cell(@line,'B')}")
			
			template.set(@line, 'C', "Topart - Special Products")
			template.set(@line, 'D', "simple")
			template.set(@line, 'E', "Collections/Oscar Night")
			template.set(@line, 'F', "Root Category")
			template.set(@line, 'G', "base")
			
			template.set(@line, 'H', "#{source.cell(@line,'S')}")
			template.set(@line, 'I', "#{source.cell(@line,'T')}")
			template.set(@line, 'J', "#{source.cell(@line,'U')}")
			template.set(@line, 'K', "#{source.cell(@line,'V')}")
			
			template.set(@line, 'L', "#{source.cell(@line,'I')}")
			template.set(@line, 'M', "#{source.cell(@line,'J')}")
			
			# Color: Look into "keywords" i.e. "AK", and search for colors...
			# ...and add each color to the same column but one @line below
			@color_set = Array["red", "orange", "yellow", "green", "blue", "purple", "pink", "brown", "black/white", "black", "white"]
			@color_count = 0;
			0.upto(@color_set.length) do |n|
				if "#{source.cell(@line,'AK')}".include? "#{@color_set[n]}"
					#@line = @line + @color_count
					template.set(@line + @color_count, 'N', "#{@color_set[n]}")
					@color_count = @color_count + 1
				end
			end
			
			# Color Tone
			template.set(@line, 'O', "#{source.cell(@line,'AH')}")
			
			template.set(@line, 'AL', "0")
			template.set(@line, 'BE', "Use config")
			template.set(@line, 'BF', "Use config")
			template.set(@line, 'BL', "Block after Info Column")
			template.set(@line, 'CJ', "1")
			
			template.set(@line, 'BZ', "1")
			template.set(@line, 'CV', "1")
			template.set(@line, 'CU', "4")
			template.set(@line, 'CN', "0")
	  
			@line = @line + @color_count
			@line = @line + 1
		end
		
		template.to_csv("filled_layered_template.csv")
 
		respond_to do |format|
			format.html # index.html.erb
		end

	end

end
