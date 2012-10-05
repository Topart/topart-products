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
		
			# Sku
			template.set(@line, 'A', "#{source.cell(@line,'B')}")
			
			template.set(@line, 'C', "Topart - Special Products")
			template.set(@line, 'D', "simple")
			template.set(@line, 'E', "Collections/Oscar Night")
			template.set(@line, 'F', "Root Category")
			template.set(@line, 'G', "base")
			
			# Alt Size 1, Alt Size 2, Alt Size 3, Alt Size 4
			template.set(@line, 'H', "#{source.cell(@line,'S')}")
			template.set(@line, 'I', "#{source.cell(@line,'T')}")
			template.set(@line, 'J', "#{source.cell(@line,'U')}")
			template.set(@line, 'K', "#{source.cell(@line,'V')}")
			
			# Artist First Name, Artist Last Name
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
			
			# Created By
			template.set(@line, 'S', "#{source.cell(@line,'AM')}")
			
			# Date Created, Date Expired, Date Modified
			template.set(@line, 'X', "#{source.cell(@line,'AL')}")
			template.set(@line, 'Y', "#{source.cell(@line,'AP')}")
			template.set(@line, 'Z', "#{source.cell(@line,'AN')}")
			
			template.set(@line, 'AA', "#{source.cell(@line,'K')}")
			
			# Discontinued: it affects the product STATUS
			# If Yes, STATUS = 2, else STATUS = 1
			if "#{source.cell(@line,'AR')}" == "Y"
				template.set(@line, 'AB', "Yes")

				# Status
				template.set(@line, 'CA', "2")
			end
			if "#{source.cell(@line,'AR')}" == "N"
				template.set(@line, 'AB', "No")
				
				# Status
				template.set(@line, 'CA', "1")
			end	
			
			# Do Not Display: it affects the product VISIBILITY
			# If Yes, visibility = 1 else visibility = 4
			if "#{source.cell(@line,'AX')}" == "1"
				template.set(@line, 'AC', "Yes")
				
				# Visibility
				template.set(@line, 'CI', "1")
			end
			if "#{source.cell(@line,'AX')}" == "0"
				template.set(@line, 'AC', "No")
				
				# Visibility
				template.set(@line, 'CI', "4")
			end
			
			
			
			# Embellishments
			@embellishments_count = 0
			if "#{source.cell(@line,'AB')}" == "Y"
				template.set(@line + @embellishments_count, 'AD', "Metallic")
				@embellishments_count = @embellishments_count + 1
			end
			if "#{source.cell(@line,'AA')}" == "Y"
				template.set(@line + @embellishments_count, 'AD', "Foil")
				@embellishments_count = @embellishments_count + 1
			end
			if "#{source.cell(@line,'X')}" == "Y"
				template.set(@line + @embellishments_count, 'AD', "Serigraph")
				@embellishments_count = @embellishments_count + 1
			end
			if "#{source.cell(@line,'Y')}" == "Y"
				template.set(@line + @embellishments_count, 'AD', "Embossed")
				@embellishments_count = @embellishments_count + 1
			end
			
			# Format: parse the width and the height, compute a rounded ratio and then the format
			@page_size_cm = "#{source.cell(@line,'M')}"
			
			@width = @page_size_cm.gsub(/ x .[0-9]/, "").to_f
			@height = @page_size_cm.gsub(/.[0-9] x /, "").to_f
			
			@ratio = @width / @height
			if @ratio > 0.9 && @ratio < 1.1
				template.set(@line, 'AF', "square")
			end
			if @ratio > 0.4 && @ratio < 0.7
				template.set(@line, 'AF', "portrait")
			end
			if @ratio > 1.5 && @ratio < 2.1
				template.set(@line, 'AF', "landscape")
			end
			if @ratio > 0.24 && @ratio < 0.4
				template.set(@line, 'AF', "panel")
			end
			if @ratio > 2.8 && @ratio < 4.2
				template.set(@line, 'AF', "panorama")
			end
			
			# Image
			template.set(@line, 'AK', "#{source.cell(@line,'Q')}")
			template.set(@line, 'AL', "#{source.cell(@line,'K')}")
			template.set(@line, 'AM', "#{source.cell(@line,'O')}")
			template.set(@line, 'AN', "#{source.cell(@line,'P')}")
			
			# Keywords
			template.set(@line, 'AO', "#{source.cell(@line,'AK')}")
			
			# LLC Stock
			template.set(@line, 'AP', "#{source.cell(@line,'G')}")
			
			# Title
			template.set(@line, 'BC', "#{source.cell(@line,'K')}")
			
			# No canvas
			if "#{source.cell(@line,'AV')}" == "Y"
				template.set(@line, 'BF', "Yes")
			end	
			if "#{source.cell(@line,'AV')}" == "N"
				template.set(@line, 'BF', "No")
			end

			# Oversize
			if "#{source.cell(@line,'W')}" == "Y"
				template.set(@line, 'BH', "Yes")
			end	
			if "#{source.cell(@line,'W')}" == "N"
				template.set(@line, 'BH', "No")
			end

			# Page
			template.set(@line, 'BI', "#{source.cell(@line,'AE')}")
			
			# Page Size CM
			template.set(@line, 'BK', "#{source.cell(@line,'M')}")
			
			# Page Size Inches
			template.set(@line, 'BM', "#{source.cell(@line,'N')}")
			
			# POD
			if "#{source.cell(@line,'AS')}" == "Y"
				template.set(@line, 'BN', "Yes")
			end	
			if "#{source.cell(@line,'AS')}" == "N"
				template.set(@line, 'BN', "No")
			end
			
			# POD Only
			if "#{source.cell(@line,'AU')}" == "Y"
				template.set(@line, 'BO', "Yes")
			end	
			if "#{source.cell(@line,'AU')}" == "N"
				template.set(@line, 'BO', "No")
			end
			
			# Price
			template.set(@line, 'BP', "#{source.cell(@line,'L')}")
			
			
			
			template.set(@line, 'BA', "Use config")
			template.set(@line, 'BB', "Use config")
			template.set(@line, 'BG', "Block after Info Column")
			
			# Required Options
			template.set(@line, 'BQ', "0")
			
			# Short Description
			template.set(@line, 'BR', "K")
			
			# Size Category
			template.set(@line, 'BT', "#{source.cell(@line,'AG')}")
			
			# Small Image
			template.set(@line, 'BU', "#{source.cell(@line,'Q')}")
			
			# Small Image Label
			template.set(@line, 'BV', "#{source.cell(@line,'K')}")
			
			# SRL Stock
			template.set(@line, 'BZ', "#{source.cell(@line,'H')}")
			
			# Tax Class Id
			template.set(@line, 'CB', "0")
			
			# Thumbnail
			template.set(@line, 'CC', "#{source.cell(@line,'Q')}")
			
			# Thumbnail Label
			template.set(@line, 'CD', "#{source.cell(@line,'K')}")
			
			# Title
			template.set(@line, 'BV', "#{source.cell(@line,'K')}")
			
			# URL Key
			@url_key = "#{source.cell(@line,'K')}";
			@url_key = @url_key.gsub(/ /, "-")
			template.set(@line, 'CG', "#{@url_key}")
			
			# URL Path
			@url_path = @url_key + ".html"
			template.set(@line, 'CH', "#{@url_path}")
			
			# Weight
			template.set(@line, 'CK', "0")
			
			# Weight
			template.set(@line, 'CL', "0")
			
			# Weight
			template.set(@line, 'CM', "0")
			
			# Weight
			template.set(@line, 'CN', "0")
			
			# Weight
			template.set(@line, 'CO', "0")
			
			# Weight
			template.set(@line, 'CP', "0")
			
			# Weight
			template.set(@line, 'CQ', "0")
			
			# Weight
			template.set(@line, 'CR', "0")
			
			# Weight
			template.set(@line, 'CS', "0")
			
			# Weight
			template.set(@line, 'CT', "0")
			
			# Weight
			template.set(@line, 'CU', "0")
			
			# Weight
			template.set(@line, 'CW', "0")
			
			# Weight
			template.set(@line, 'CX', "0")
			
			# Weight
			template.set(@line, 'CY', "0")
			
			# Weight
			template.set(@line, 'CZ', "1")
			
			# Weight
			template.set(@line, 'DA', "0")
			
			# Weight
			template.set(@line, 'DB', "0")
			
			# Weight
			template.set(@line, 'DC', "1")
			
			# Weight
			template.set(@line, 'DD', "0")
			
			# Weight
			template.set(@line, 'DE', "0")
			
			
	  
			@line = @line + @color_count
			@line = @line + 1
		end
		
		template.to_csv("filled_layered_template.csv")
 
		respond_to do |format|
			format.html # index.html.erb
		end

	end

end
