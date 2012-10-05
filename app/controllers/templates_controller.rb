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
		
		# Color set
		@color_set = Array["red", "orange", "yellow", "green", "blue", "purple", "pink", "brown", "black/white", "black", "white"]
		
		# Fill every line in the template file up with
		# the right value taken from the source input file
		
		@destination_line = 2
		2.upto(source.last_row) do |source_line|
		
			# Sku
			template.set(@destination_line, 'A', "#{source.cell(source_line,'B')}")
			
			template.set(@destination_line, 'C', "Topart - Special Products")
			template.set(@destination_line, 'D', "simple")
			template.set(@destination_line, 'E', "Collections/Oscar Night")
			template.set(@destination_line, 'F', "Root Category")
			template.set(@destination_line, 'G', "base")
			
			# Alt Size 1, Alt Size 2, Alt Size 3, Alt Size 4
			template.set(@destination_line, 'H', "#{source.cell(source_line,'S')}")
			template.set(@destination_line, 'I', "#{source.cell(source_line,'T')}")
			template.set(@destination_line, 'J', "#{source.cell(source_line,'U')}")
			template.set(@destination_line, 'K', "#{source.cell(source_line,'V')}")
			
			# Artist First Name, Artist Last Name
			template.set(@destination_line, 'L', "#{source.cell(source_line,'I')}")
			template.set(@destination_line, 'M', "#{source.cell(source_line,'J')}")
			
			# Color: Look into "keywords" i.e. "AK", and search for colors...
			# ...and add each color to the same column but one @destination_line below
			@color_count = 0;
			0.upto(@color_set.length) do |n|
				if "#{source.cell(source_line,'AK')}".include? "#{@color_set[n]}"
					template.set(@destination_line + @color_count, 'N', "#{@color_set[n]}")
					@color_count = @color_count + 1
				end
			end
			
			# Color Tone
			template.set(@destination_line, 'O', "#{source.cell(source_line,'AH')}")
			
			# Created By
			template.set(@destination_line, 'S', "#{source.cell(source_line,'AM')}")
			
			# Date Created, Date Expired, Date Modified
			template.set(@destination_line, 'X', "#{source.cell(source_line,'AL')}")
			template.set(@destination_line, 'Y', "#{source.cell(source_line,'AP')}")
			template.set(@destination_line, 'Z', "#{source.cell(source_line,'AN')}")
			
			template.set(@destination_line, 'AA', "#{source.cell(source_line,'K')}")
			
			# Discontinued: it affects the product STATUS
			# If Yes, STATUS = 2, else STATUS = 1
			if "#{source.cell(source_line,'AR')}" == "Y"
				template.set(@destination_line, 'AB', "Yes")

				# Status
				template.set(@destination_line, 'CA', "2")
			end
			if "#{source.cell(source_line,'AR')}" == "N"
				template.set(@destination_line, 'AB', "No")
				
				# Status
				template.set(@destination_line, 'CA', "1")
			end	
			
			# Do Not Display: it affects the product VISIBILITY
			# If Yes, visibility = 1 else visibility = 4
			if "#{source.cell(source_line,'AX')}" == "1"
				template.set(@destination_line, 'AC', "Yes")
				
				# Visibility
				template.set(@destination_line, 'CI', "1")
			end
			if "#{source.cell(source_line,'AX')}" == "0"
				template.set(@destination_line, 'AC', "No")
				
				# Visibility
				template.set(@destination_line, 'CI', "4")
			end
			
			
			
			# Embellishments
			@embellishments_count = 0
			if "#{source.cell(source_line,'AB')}" == "Y"
				template.set(@destination_line + @embellishments_count, 'AD', "Metallic")
				@embellishments_count = @embellishments_count + 1
			end
			if "#{source.cell(source_line,'AA')}" == "Y"
				template.set(@destination_line + @embellishments_count, 'AD', "Foil")
				@embellishments_count = @embellishments_count + 1
			end
			if "#{source.cell(source_line,'X')}" == "Y"
				template.set(@destination_line + @embellishments_count, 'AD', "Serigraph")
				@embellishments_count = @embellishments_count + 1
			end
			if "#{source.cell(source_line,'Y')}" == "Y"
				template.set(@destination_line + @embellishments_count, 'AD', "Embossed")
				@embellishments_count = @embellishments_count + 1
			end
			
			# Format: parse the width and the height, compute a rounded ratio and then the format
			@page_size_cm = "#{source.cell(source_line,'M')}"
			
			@width = @page_size_cm.gsub(/ x .[0-9]/, "").to_f
			@height = @page_size_cm.gsub(/.[0-9] x /, "").to_f
			
			@ratio = @width / @height
			if @ratio > 0.9 && @ratio < 1.1
				template.set(@destination_line, 'AF', "square")
			end
			if @ratio > 0.4 && @ratio < 0.7
				template.set(@destination_line, 'AF', "portrait")
			end
			if @ratio > 1.5 && @ratio < 2.1
				template.set(@destination_line, 'AF', "landscape")
			end
			if @ratio > 0.24 && @ratio < 0.4
				template.set(@destination_line, 'AF', "panel")
			end
			if @ratio > 2.8 && @ratio < 4.2
				template.set(@destination_line, 'AF', "panorama")
			end
			
			# Image
			template.set(@destination_line, 'AK', "#{source.cell(source_line,'Q')}")
			template.set(@destination_line, 'AL', "#{source.cell(source_line,'K')}")
			template.set(@destination_line, 'AM', "#{source.cell(source_line,'O')}")
			template.set(@destination_line, 'AN', "#{source.cell(source_line,'P')}")
			
			# Keywords
			template.set(@destination_line, 'AO', "#{source.cell(source_line,'AK')}")
			
			# LLC Stock
			template.set(@destination_line, 'AP', "#{source.cell(source_line,'G')}")
			
			# Title
			template.set(@destination_line, 'BC', "#{source.cell(source_line,'K')}")
			
			# No canvas
			if "#{source.cell(source_line,'AV')}" == "Y"
				template.set(@destination_line, 'BF', "Yes")
			end	
			if "#{source.cell(source_line,'AV')}" == "N"
				template.set(@destination_line, 'BF', "No")
			end

			# Oversize
			if "#{source.cell(source_line,'W')}" == "Y"
				template.set(@destination_line, 'BH', "Yes")
			end	
			if "#{source.cell(source_line,'W')}" == "N"
				template.set(@destination_line, 'BH', "No")
			end

			# Page
			template.set(@destination_line, 'BI', "#{source.cell(source_line,'AE')}")
			
			# Page Size CM
			template.set(@destination_line, 'BK', "#{source.cell(source_line,'M')}")
			
			# Page Size Inches
			template.set(@destination_line, 'BM', "#{source.cell(source_line,'N')}")
			
			# POD
			if "#{source.cell(source_line,'AS')}" == "Y"
				template.set(@destination_line, 'BN', "Yes")
			end	
			if "#{source.cell(source_line,'AS')}" == "N"
				template.set(@destination_line, 'BN', "No")
			end
			
			# POD Only
			if "#{source.cell(source_line,'AU')}" == "Y"
				template.set(@destination_line, 'BO', "Yes")
			end	
			if "#{source.cell(source_line,'AU')}" == "N"
				template.set(@destination_line, 'BO', "No")
			end
			
			# Price
			template.set(@destination_line, 'BP', "#{source.cell(source_line,'L')}")
			
			
			
			template.set(@destination_line, 'BA', "Use config")
			template.set(@destination_line, 'BB', "Use config")
			template.set(@destination_line, 'BG', "Block after Info Column")
			
			# Required Options
			template.set(@destination_line, 'BQ', "0")
			
			# Short Description
			template.set(@destination_line, 'BR', "K")
			
			# Size Category
			template.set(@destination_line, 'BT', "#{source.cell(source_line,'AG')}")
			
			# Small Image
			template.set(@destination_line, 'BU', "#{source.cell(source_line,'Q')}")
			
			# Small Image Label
			template.set(@destination_line, 'BV', "#{source.cell(source_line,'K')}")
			
			# SRL Stock
			template.set(@destination_line, 'BZ', "#{source.cell(source_line,'H')}")
			
			# Tax Class Id
			template.set(@destination_line, 'CB', "0")
			
			# Thumbnail
			template.set(@destination_line, 'CC', "#{source.cell(source_line,'Q')}")
			
			# Thumbnail Label
			template.set(@destination_line, 'CD', "#{source.cell(source_line,'K')}")
			
			# Title
			template.set(@destination_line, 'BV', "#{source.cell(source_line,'K')}")
			
			# URL Key
			@url_key = "#{source.cell(source_line,'K')}";
			@url_key = @url_key.gsub(/ /, "-")
			template.set(@destination_line, 'CG', "#{@url_key}")
			
			# URL Path
			@url_path = @url_key + ".html"
			template.set(@destination_line, 'CH', "#{@url_path}")
			
			# Weight
			template.set(@destination_line, 'CK', "0")
			
			# Weight
			template.set(@destination_line, 'CL', "0")
			
			# Weight
			template.set(@destination_line, 'CM', "0")
			
			# Weight
			template.set(@destination_line, 'CN', "0")
			
			# Weight
			template.set(@destination_line, 'CO', "0")
			
			# Weight
			template.set(@destination_line, 'CP', "0")
			
			# Weight
			template.set(@destination_line, 'CQ', "0")
			
			# Weight
			template.set(@destination_line, 'CR', "0")
			
			# Weight
			template.set(@destination_line, 'CS', "0")
			
			# Weight
			template.set(@destination_line, 'CT', "0")
			
			# Weight
			template.set(@destination_line, 'CU', "0")
			
			# Weight
			template.set(@destination_line, 'CW', "0")
			
			# Weight
			template.set(@destination_line, 'CX', "0")
			
			# Weight
			template.set(@destination_line, 'CY', "0")
			
			# Weight
			template.set(@destination_line, 'CZ', "1")
			
			# Weight
			template.set(@destination_line, 'DA', "0")
			
			# Weight
			template.set(@destination_line, 'DB', "0")
			
			# Weight
			template.set(@destination_line, 'DC', "1")
			
			# Weight
			template.set(@destination_line, 'DD', "0")
			
			# Weight
			template.set(@destination_line, 'DE', "0")
			
			
			
			# Compute the maximum count among all the multi select options
			# then add it to the destination line count for the next product to be written
			
			@multi_select_options = Array.new
			@multi_select_options << @color_count << @embellishments_count
			@max_count =  @multi_select_options.max
			
			# Increase the destination line to the correct number
			@destination_line = @destination_line + @max_count
			@destination_line = @destination_line + 1
		end
		
		# Finally, fill the template
		template.to_csv("filled_layered_template.csv")
 
		# Accessing this view launch the service automatically
		respond_to do |format|
			format.html # index.html.erb
		end

	end

end
