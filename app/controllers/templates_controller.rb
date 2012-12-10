require 'rubygems'
require 'roo'

class TemplatesController < ApplicationController

	# GET /generate_template
	# GET /generate_template.json
	def index
 
		# Load the source Excel file, with all the special products info
		#source = Excel.new("http://beta.topart.com/csv/Template_2012_11_28/source.xls")
		source = Excel.new("Template_2012_11_28/source.xls")
		source.default_sheet = source.sheets.first
		
		# Load the Magento template, which is in Open Office format
		#template = Openoffice.new("http://beta.topart.com/csv/Template_2012_11_28/template.ods")
		template = Openoffice.new("Template_2012_11_28/template.ods")
		template.default_sheet = template.sheets.first
		
		# Color set
		@color_set = Array["red", "orange", "yellow", "green", "blue", "purple", "pink", "brown", "black/white", "black", "white"]

		# Categories list
		source_categories = Excel.new("Template_2012_11_28/category list for website.xls")
		source_categories.default_sheet = source_categories.sheets.first

		

		# Fill every line in the template file up with
		# the right value taken from the source input file
		
		@destination_line = 2
		2.upto(source.last_row) do |source_line|
		#2.upto(100) do |source_line|
		
			# Sku: insert a "_" after the last character in the string
			original_sku = "#{source.cell(source_line,'A')}"
			#sku = original_sku.sub!(match, '_')
			template.set(@destination_line, 'A', original_sku)
			
			template.set(@destination_line, 'C', "Topart - Products")
			template.set(@destination_line, 'D', "simple")

			@collections_count = 0

			# Artist Focus: look for artists names
			if ( "#{source.cell(source_line,'C')}".downcase.strip == "chris donovan" or "#{source.cell(source_line,'C')}".downcase.strip == "luke wilson" or "#{source.cell(source_line,'C')}".downcase.strip == "erin lange" or "#{source.cell(source_line,'C')}".downcase.strip == "gregory williams" or "#{source.cell(source_line,'C')}".downcase.strip == "john seba" or "#{source.cell(source_line,'C')}".downcase.strip == "mike klung" or "#{source.cell(source_line,'C')}".downcase.strip == "alex edwards")

				template.set(@destination_line + @collections_count, 'E', "Artist Focus/" + "#{source.cell(source_line,'C')}")
				template.set(@destination_line + @collections_count, 'F', "Root Category")

				@collections_count = @collections_count + 1

			end

			# Check if any keyword matches any category name. 
			# If there is a match, add the category name to the corresponding product row.
			2.upto(source_categories.last_row) do |source_categories_line|
			#3.upto(3) do |source_categories_line|

			# Discard the "delete" categories
				if ( "#{source_categories.cell(source_categories_line,'C')}".downcase.strip != "delete" )

					# Check for set intersection between keywords and category names keywords
					if ( ( ("#{source.cell(source_line,'AT')}".downcase.split(";")) & ("#{source_categories.cell(source_categories_line,'B')}".downcase.split(",") ) ).any? )
				
						template.set(@destination_line + @collections_count, 'E', "Browse Art/" + "#{source_categories.cell(source_categories_line,'A')}".strip)
						template.set(@destination_line + @collections_count, 'F', "Root Category")

						@collections_count = @collections_count + 1

					end
				end
			end

			# Featured Collections
			### Empty for now

			# Oversize Variety
			if ( "#{source.cell(source_line,'N')}" == "Y")

				template.set(@destination_line + @collections_count, 'E', "Collections/Oversize Variety")
				template.set(@destination_line + @collections_count, 'F', "Root Category")

				@collections_count = @collections_count + 1

			end

			# Abastract Geometry
			if "#{source.cell(source_line,'AT')}".downcase.include? "industrial" or "#{source.cell(source_line,'AT')}".downcase.include? "geometric"

				template.set(@destination_line + @collections_count, 'E', "Collections/Abstract Geometry")
				template.set(@destination_line + @collections_count, 'F', "Root Category")

				@collections_count = @collections_count + 1

			end
			

			# Urban Industrial
			if "#{source.cell(source_line,'AT')}".downcase.include? "industrial"

				template.set(@destination_line + @collections_count, 'E', "Collections/Urban Industrial")
				template.set(@destination_line + @collections_count, 'F', "Root Category")

				@collections_count = @collections_count + 1

			end

			# Gustav Klimt
			if "#{source.cell(source_line,'AT')}".downcase.include? "klimt"

				template.set(@destination_line + @collections_count, 'E', "Collections/Gustav Klimt-150th Anniversary")
				template.set(@destination_line + @collections_count, 'F', "Root Category")

				@collections_count = @collections_count + 1

			end

			template.set(@destination_line, 'G', "base")
			
			# Alt Size 1, Alt Size 2, Alt Size 3, Alt Size 4
			template.set(@destination_line, 'H', "#{source.cell(source_line,'J')}")
			template.set(@destination_line, 'I', "#{source.cell(source_line,'K')}")
			template.set(@destination_line, 'J', "#{source.cell(source_line,'L')}")
			template.set(@destination_line, 'K', "#{source.cell(source_line,'M')}")
			
			# Color: Look into "keywords" and search for colors...
			# ...and add each color to the same column but on one @destination_line below
			@color_count = 0;
			0.upto(@color_set.length) do |n|
				if "#{source.cell(source_line,'AT')}".downcase.include? "#{@color_set[n]}"
					template.set(@destination_line + @color_count, 'L', "#{@color_set[n]}")
					@color_count = @color_count + 1
				end
			end
			
			# Color Tone
			#template.set(@destination_line, 'M', "#{source.cell(source_line,'AH')}")
			
			#Description
			template.set(@destination_line, 'Y', "#{source.cell(source_line,'B')}")


			# Embellishments
			@embellishments_count = 0
			if "#{source.cell(source_line,'S')}" == "Y"
				template.set(@destination_line + @embellishments_count, 'AA', "Metallic")
				@embellishments_count = @embellishments_count + 1
			end
			if "#{source.cell(source_line,'R')}" == "Y"
				template.set(@destination_line + @embellishments_count, 'AA', "Foil")
				@embellishments_count = @embellishments_count + 1
			end
			if "#{source.cell(source_line,'O')}" == "Y"
				template.set(@destination_line + @embellishments_count, 'AA', "Serigraph")
				@embellishments_count = @embellishments_count + 1
			end
			if "#{source.cell(source_line,'P')}" == "Y"
				template.set(@destination_line + @embellishments_count, 'AA', "Embossed")
				@embellishments_count = @embellishments_count + 1
			end


			#Enable Google Checkout
			template.set(@destination_line, 'AB', "1")


			# Orientation: parse the width and the height, compute a rounded ratio and then the format
			#@image_size_cm = "#{source.cell(source_line,'M')}"
			
			#@width = @page_size_cm.gsub(/ x .[0-9]/, "").to_f
			#@height = @page_size_cm.gsub(/.[0-9] x /, "").to_f
			
			#@ratio = @width / @height
			#if @ratio > 0.9 && @ratio < 1.1
			#	template.set(@destination_line, 'AF', "square")
			#end
			#if @ratio > 0.4 && @ratio < 0.7
			#	template.set(@destination_line, 'AF', "portrait")
			#end
			#if @ratio > 1.5 && @ratio < 2.1
			#	template.set(@destination_line, 'AF', "landscape")
			#end
			
			#if @ratio > 0.24 && @ratio < 0.4
			#	template.set(@destination_line, 'AF', "panel")
			#end
			#if @ratio > 2.8 && @ratio < 4.2
			#	template.set(@destination_line, 'AF', "panorama")
			#end

			#Orientation: get it directly from it corresponding column
			template.set(@destination_line, 'AC', "#{source.cell(source_line,'T')}")


			#has_options
			if "#{source.cell(source_line,'A')}" =~ /DG$/ 
				template.set(@destination_line, 'AF', "1")
			else
				template.set(@destination_line, 'AF', "0")
			end

			# Image size cm
			template.set(@destination_line, 'AI', "#{source.cell(source_line,'H')}")

			# Image size inches
			template.set(@destination_line, 'AJ', "#{source.cell(source_line,'I')}")

			#Keywords
			template.set(@destination_line, 'AK', "#{source.cell(source_line,'AT')}".downcase)


			#Meta Description
			template.set(@destination_line, 'AN', "#{source.cell(source_line,'B')}")

			#Meta Kewyord
			template.set(@destination_line, 'AO', "#{source.cell(source_line,'AT')}")

			#Meta title
			template.set(@destination_line, 'AP', "#{source.cell(source_line,'B')}")


			#msrp_display_actual_price_type
			template.set(@destination_line, 'AT', "Use config")
			
			#msrp_enabled
			template.set(@destination_line, 'AU', "Use config")

			#Name
			template.set(@destination_line, 'AV', "#{source.cell(source_line,'B')}")

			#options_container
			template.set(@destination_line, 'AY', "Block after Info Column")

			#Oversize
			if "#{source.cell(source_line,'N')}" == "Y"
				template.set(@destination_line, 'AZ', "#{source.cell(source_line,'Yes')}")
			else
				template.set(@destination_line, 'AZ', "#{source.cell(source_line,'No')}")
			end

			#Paper size cm
			template.set(@destination_line, 'BC', "#{source.cell(source_line,'F')}")

			#Paper size inches
			template.set(@destination_line, 'BE', "#{source.cell(source_line,'G')}")

			#A4POD
			if "#{source.cell(source_line,'V')}" == "Y"
				template.set(@destination_line, 'BF', "Yes")
			else
				template.set(@destination_line, 'BF', "No")
			end

			#Price
			template.set(@destination_line, 'BG', "#{source.cell(source_line,'E')}")


			#required_options
			if "#{source.cell(source_line,'A')}" =~ /DG$/ 
				template.set(@destination_line, 'BH', "1")
			else
				template.set(@destination_line, 'BH', "0")
			end

			#Short description
			template.set(@destination_line, 'BI', "#{source.cell(source_line,'B')}")


			#Size category: for posters
			@image_size_cm = "#{source.cell(source_line,'H')}"
			
			@width = @image_size_cm.gsub(/ x .[0-9]/, "")
			@height = @image_size_cm.gsub(/.[0-9] x /, "")

			#Convert UI to inches to have a consistent comparison with the spreadsheet
			@ui = ( (@width.to_i + @height.to_i) / 2.54).to_i;

			if (@ui != 0)

				if @ui < 40 
					template.set(@destination_line, 'BK', "Petite")
				end

				if @ui >= 40 and @ui <  50
					template.set(@destination_line, 'BK', "Small")
				end

				if @ui >= 50 and @ui < 60 
					template.set(@destination_line, 'BK', "Medium")
				end

				if @ui >= 60 and @ui < 70
					template.set(@destination_line, 'BK', "Large")
				end

				if @ui >= 70   
					template.set(@destination_line, 'BK', "Oversize")
				end

			end

			#Status: enabled (1), disabled (2)
			if "#{source.cell(source_line,'A')}" =~ /DG$/ 
				template.set(@destination_line, 'BQ', "2")
			else
				template.set(@destination_line, 'BQ', "1")
			end

			#Tax class ID
			template.set(@destination_line, 'BR', "2")

			#total_quantity_on_hand
			if "#{source.cell(source_line,'A')}" =~ /DG$/ 
				template.set(@destination_line, 'BV', "0".to_i)
			else
				template.set(@destination_line, 'BV', "#{source.cell(source_line,'AE')}".to_i)
			end

			#udf_anycustom
			if "#{source.cell(source_line,'AR')}" == "Y"
				template.set(@destination_line, 'BW', "Yes")
			else
				template.set(@destination_line, 'BW', "No")
			end

			#Artist name
			template.set(@destination_line, 'BX', "#{source.cell(source_line,'C')}")

			#Copyright
			if "#{source.cell(source_line,'AP')}" == "Y"
				template.set(@destination_line, 'BY', "Yes")
			else
				template.set(@destination_line, 'BY', "No")
			end

			#udf_crline
			template.set(@destination_line, 'BZ', "#{source.cell(source_line,'AQ')}")

			#udf_dnd
			if "#{source.cell(source_line,'X')}" == "Y"
				template.set(@destination_line, 'CB', "Yes")
			else
				template.set(@destination_line, 'CB', "No")
			end

			#udf_embellished
			if "#{source.cell(source_line,'AG')}" == "Y"
				template.set(@destination_line, 'CC', "Yes")
			else
				template.set(@destination_line, 'CC', "No")
			end

			#udf_framed
			if "#{source.cell(source_line,'AH')}" == "Y"
				template.set(@destination_line, 'CD', "Yes")
			else
				template.set(@destination_line, 'CD', "No")
			end

			#udf_imsource
			template.set(@destination_line, 'CE', "#{source.cell(source_line,'Y')}")

			#udf_limited
			if "#{source.cell(source_line,'AO')}" == "Y"
				template.set(@destination_line, 'CF', "Yes")
			else
				template.set(@destination_line, 'CF', "No")
			end

			#udf_maxsf
			template.set(@destination_line, 'CG', "0")

			#udf_new
			if "#{source.cell(source_line,'U')}" == "Y"
				template.set(@destination_line, 'CH', "Yes")
			else
				template.set(@destination_line, 'CH', "No")
			end

			#udf_osdp
			if "#{source.cell(source_line,'AN')}" == "Y"
				template.set(@destination_line, 'CI', "Yes")
			else
				template.set(@destination_line, 'CI', "No")
			end

			#udf_pricecorde
			template.set(@destination_line, 'CJ', "#{source.cell(source_line,'D')}")

			#udf_ratiocode
			time = "#{source.cell(source_line,'W')}".to_i
			hours = time/3600
			minutes = (time/60 - hours * 60)

			ratio_code = hours.to_s << ":" << minutes.to_s
			template.set(@destination_line, 'CK', ratio_code)

			#udf_tar
			if "#{source.cell(source_line,'Z')}" == "Y"
				template.set(@destination_line, 'CL', "Yes")
			else
				template.set(@destination_line, 'CL', "No")
			end

			#URL Key, with the SKU as suffix to keep it unique among products
			template.set(@destination_line, 'CN', "#{source.cell(source_line,'B')}".gsub(/[ ]/, '-')  << "-" << "#{source.cell(source_line,'A')}")

			#Visibility
			template.set(@destination_line, 'CP', "4")

			#Weight
			template.set(@destination_line, 'CQ', "1")

			#Qty
			if "#{source.cell(source_line,'A')}" =~ /DG$/ 
				template.set(@destination_line, 'CR', "0")
			else
				template.set(@destination_line, 'CR', "#{source.cell(source_line,'AE')}")
			end

			#Min qty
			template.set(@destination_line, 'CS', "0")

			#use_config_min_qty
			template.set(@destination_line, 'CT', "1")

			#is_qty_decimal
			template.set(@destination_line, 'CU', "0")

			#backorders
			template.set(@destination_line, 'CV', "0")

			#use_config_backorders
			template.set(@destination_line, 'CW', "1")

			#min_sale_qty
			template.set(@destination_line, 'CX', "1")

			#use_config_min_sale_qty
			template.set(@destination_line, 'CY', "1")

			#max_sale_qty
			template.set(@destination_line, 'CZ', "0")

			#use_config_max_sale_qty
			template.set(@destination_line, 'DA', "1")

			#is_in_stock
			if "#{source.cell(source_line,'A')}" =~ /DG$/
				template.set(@destination_line, 'DB', "1")
			else
				template.set(@destination_line, 'DB', "0")
			end

			#use_config_notify_stock_qty
			if "#{source.cell(source_line,'A')}" =~ /DG$/
				template.set(@destination_line, 'DD', "1")
			else
				template.set(@destination_line, 'DD', "0")
			end

			#manage_stock
			if "#{source.cell(source_line,'A')}" =~ /DG$/
				template.set(@destination_line, 'DE', "1")
			else
				template.set(@destination_line, 'DE', "0")
			end

			#use_config_manage_stock
			if "#{source.cell(source_line,'A')}" =~ /DG$/
				template.set(@destination_line, 'DF', "1")
			else
				template.set(@destination_line, 'DF', "0")
			end

			#stock_status_changed_auto
			template.set(@destination_line, 'DG', "0")

			#use_config_qty_increments
			template.set(@destination_line, 'DH', "1")

			#qty_increments
			template.set(@destination_line, 'DI', "0")

			#use_config_enable_qty_inc
			template.set(@destination_line, 'DJ', "1")

			#enable_qty_increments
			template.set(@destination_line, 'DK', "0")

			#is_decimal_divided
			template.set(@destination_line, 'DL', "0")



			########## Custom options columns ##########
			if "#{source.cell(source_line,'A')}" =~ /DG$/

				#############SIZE#############

				#_custom_option_store
				#template.set(@destination_line, 'EH', "default")
				#_custom_option_type
				template.set(@destination_line, 'EI', "radio")
				#_custom_option_title
				template.set(@destination_line, 'EJ', "Size")
				#_custom_option_is_required
				template.set(@destination_line, 'EK', "1")
				#_custom_option_max_characters
				template.set(@destination_line, 'EN', "0")
				#_custom_option_sort_order
				template.set(@destination_line, 'EO', "3")
				

				#_custom_option_row_title
				template.set(@destination_line, 'EP', "Petite")
				#_custom_option_row_price
				template.set(@destination_line, 'EQ', "0")
				#_custom_option_row_sku
				template.set(@destination_line, 'ER', "size_petite")
				#_custom_option_row_sort
				template.set(@destination_line, 'ES', "0")

				@destination_line = @destination_line + 1

				#_custom_option_row_title
				template.set(@destination_line, 'EP', "Small")
				#_custom_option_row_price
				template.set(@destination_line, 'EQ', "0")
				#_custom_option_row_sku
				template.set(@destination_line, 'ER', "size_small")
				#_custom_option_row_sort
				template.set(@destination_line, 'ES', "1")

				@destination_line = @destination_line + 1

				#_custom_option_row_title
				template.set(@destination_line, 'EP', "Medium")
				#_custom_option_row_price
				template.set(@destination_line, 'EQ', "0")
				#_custom_option_row_sku
				template.set(@destination_line, 'ER', "size_medium")
				#_custom_option_row_sort
				template.set(@destination_line, 'ES', "2")

				@destination_line = @destination_line + 1

				#_custom_option_row_title
				template.set(@destination_line, 'EP', "Large")
				#_custom_option_row_price
				template.set(@destination_line, 'EQ', "0")
				#_custom_option_row_sku
				template.set(@destination_line, 'ER', "size_large")
				#_custom_option_row_sort
				template.set(@destination_line, 'ES', "3")

				@destination_line = @destination_line + 1

				#_custom_option_row_title
				template.set(@destination_line, 'EP', "Oversize")
				#_custom_option_row_price
				template.set(@destination_line, 'EQ', "0")
				#_custom_option_row_sku
				template.set(@destination_line, 'ER', "size_oversize")
				#_custom_option_row_sort
				template.set(@destination_line, 'ES', "4")

				@destination_line = @destination_line + 1

				#_custom_option_row_title
				template.set(@destination_line, 'EP', "Custom")
				#_custom_option_row_price
				template.set(@destination_line, 'EQ', "0")
				#_custom_option_row_sku
				template.set(@destination_line, 'ER', "size_custom")
				#_custom_option_row_sort
				template.set(@destination_line, 'ES', "5")

				@destination_line = @destination_line + 1


				########### Canvas Quality ###############

				#_custom_option_store
				#template.set(@destination_line, 'EH', "default")
				#_custom_option_type
				template.set(@destination_line, 'EI', "radio")
				#_custom_option_title
				template.set(@destination_line, 'EJ', "Canvas Quality")
				#_custom_option_is_required
				template.set(@destination_line, 'EK', "0")
				#_custom_option_max_characters
				template.set(@destination_line, 'EN', "0")
				#_custom_option_sort_order
				template.set(@destination_line, 'EO', "2")
				

				#_custom_option_row_title
				template.set(@destination_line, 'EP', "Standard")
				#_custom_option_row_price
				template.set(@destination_line, 'EQ', "0")
				#_custom_option_row_sku
				template.set(@destination_line, 'ER', "canvas_standard")
				#_custom_option_row_sort
				template.set(@destination_line, 'ES', "0")

				@destination_line = @destination_line + 1

				#_custom_option_row_title
				template.set(@destination_line, 'EP', "Rag")
				#_custom_option_row_price
				template.set(@destination_line, 'EQ', "0")
				#_custom_option_row_sku
				template.set(@destination_line, 'ER', "canvas_rag")
				#_custom_option_row_sort
				template.set(@destination_line, 'ES', "1")

				@destination_line = @destination_line + 1


				########### Paper Quality ###############

				#_custom_option_store
				#template.set(@destination_line, 'EH', "default")
				#_custom_option_type
				template.set(@destination_line, 'EI', "radio")
				#_custom_option_title
				template.set(@destination_line, 'EJ', "Paper Quality")
				#_custom_option_is_required
				template.set(@destination_line, 'EK', "0")
				#_custom_option_max_characters
				template.set(@destination_line, 'EN', "0")
				#_custom_option_sort_order
				template.set(@destination_line, 'EO', "1")
				

				#_custom_option_row_title
				template.set(@destination_line, 'EP', "Standard Quality Paper")
				#_custom_option_row_price
				template.set(@destination_line, 'EQ', "0")
				#_custom_option_row_sku
				template.set(@destination_line, 'ER', "paperquality_standard")
				#_custom_option_row_sort
				template.set(@destination_line, 'ES', "0")

				@destination_line = @destination_line + 1

				#_custom_option_row_title
				template.set(@destination_line, 'EP', "Special Quality Paper")
				#_custom_option_row_price
				template.set(@destination_line, 'EQ', "0")
				#_custom_option_row_sku
				template.set(@destination_line, 'ER', "paperquality_specialpaper")
				#_custom_option_row_sort
				template.set(@destination_line, 'ES', "1")

				@destination_line = @destination_line + 1



				########### Material ###############

				#_custom_option_store
				#template.set(@destination_line, 'EH', "default")
				#_custom_option_type
				template.set(@destination_line, 'EI', "radio")
				#_custom_option_title
				template.set(@destination_line, 'EJ', "Material")
				#_custom_option_is_required
				template.set(@destination_line, 'EK', "1")
				#_custom_option_max_characters
				template.set(@destination_line, 'EN', "0")
				#_custom_option_sort_order
				template.set(@destination_line, 'EO', "0")
				

				#_custom_option_row_title
				template.set(@destination_line, 'EP', "Poster")
				#_custom_option_row_price
				template.set(@destination_line, 'EQ', "0")
				#_custom_option_row_sku
				template.set(@destination_line, 'ER', "material_poster")
				#_custom_option_row_sort
				template.set(@destination_line, 'ES', "0")

				@destination_line = @destination_line + 1

				#_custom_option_row_title
				template.set(@destination_line, 'EP', "Photopaper")
				#_custom_option_row_price
				template.set(@destination_line, 'EQ', "0")
				#_custom_option_row_sku
				template.set(@destination_line, 'ER', "material_photopaper")
				#_custom_option_row_sort
				template.set(@destination_line, 'ES', "1")

				@destination_line = @destination_line + 1

				#_custom_option_row_title
				template.set(@destination_line, 'EP', "Canvas")
				#_custom_option_row_price
				template.set(@destination_line, 'EQ', "0")
				#_custom_option_row_sku
				template.set(@destination_line, 'ER', "material_canvas")
				#_custom_option_row_sort
				template.set(@destination_line, 'ES', "2")

				@destination_line = @destination_line + 1

				#_custom_option_row_title
				template.set(@destination_line, 'EP', "Decal")
				#_custom_option_row_price
				template.set(@destination_line, 'EQ', "0")
				#_custom_option_row_sku
				template.set(@destination_line, 'ER', "material_decal")
				#_custom_option_row_sort
				template.set(@destination_line, 'ES', "3")

			end	
			
			
			
			# Compute the maximum count among all the multi select options
			# then add it to the destination line count for the next product to be written
			
			@custom_options_array_size = 0

			@multi_select_options = Array.new
			@multi_select_options << @color_count << @embellishments_count << @collections_count

			if "#{source.cell(source_line,'A')}" =~ /DG$/
				@multi_select_options << @custom_options_array_size
			end

			@max_count =  @multi_select_options.max
			
			# Increase the destination line to the correct number
			@destination_line = @destination_line + @max_count
			@destination_line = @destination_line + 1
		end
		
		# Finally, fill the template
		template.to_csv("new_inventory.csv")
 
		# Accessing this view launch the service automatically
		respond_to do |format|
			format.html # index.html.erb
		end

	end

end
