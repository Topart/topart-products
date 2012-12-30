require 'rubygems'
require 'roo'

class TemplatesController < ApplicationController

	# GET /generate_template
	# GET /generate_template.json
	def index
 
		# Load the source Excel file, with all the special products info
		#source = Excel.new("http://beta.topart.com/csv/Template_2012_11_28/source.xls")
		source = Excel.new("Template_2012_12_27/source.xls")
		source.default_sheet = source.sheets.first
		
		# Load the Magento template, which is in Open Office format
		#template = Openoffice.new("http://beta.topart.com/csv/Template_2012_11_28/template.ods")
		template = Openoffice.new("Template_2012_12_27/template.ods")
		template.default_sheet = template.sheets.first
		
		# Color set
		@color_set = Array["red", "orange", "yellow", "green", "blue", "purple", "pink", "brown", "black/white", "black", "white"]

		# Categories list
		source_categories = Excel.new("Template_2012_12_27/category list for website.xls")
		source_categories.default_sheet = source_categories.sheets.first

		# Automatically scan the template column names and store them in an associative array
		@template_dictionary = Hash.new
		"A".upto("ES") do |alphabet_character|

			@cell_content = "#{template.cell(1, alphabet_character)}"
			@template_dictionary[@cell_content] = alphabet_character
		end

		#p @template_dictionary["udf_limited"]

		# Automatically scan the source column names and store them in an associative array
		@source_dictionary = Hash.new
		"A".upto("AW") do |alphabet_character|

			@cell_content = "#{source.cell(1, alphabet_character)}"
			@source_dictionary[@cell_content] = alphabet_character
		end


		# Load the retail_material_size spreadsheet file for paper
		retail_material_size_paper = Excel.new("Template_2012_12_27/retail_material_size_treatments.xls")
		retail_material_size_paper.default_sheet = retail_material_size_paper.sheets.first

		# Load the retail_material_size spreadsheet file for canvas
		retail_material_size_canvas = Excel.new("Template_2012_12_27/retail_material_size_treatments.xls")
		retail_material_size_canvas.default_sheet = retail_material_size_canvas.sheets.second

		# Load the retail_framing_stretching_matting spreadsheet file to extract framing, stretching and matting information
		retail_framing_stretching_matting = Excel.new("Template_2012_12_27/retail_framing_stretching_matting.xls")
		retail_framing_stretching_matting.default_sheet = retail_framing_stretching_matting.sheets.first


		# MATERIAL -> PAPER
		# Automatically scan the source column names and store them in an associative array
		@retail_material_size_paper_dictionary = Hash.new
		"A".upto("S") do |alphabet_character|

			@cell_content = "#{retail_material_size_paper.cell(1, alphabet_character)}"
			@retail_material_size_paper_dictionary[@cell_content] = alphabet_character
		end

		# MATERIAL -> CANVAS
		# Automatically scan the source column names and store them in an associative array
		@retail_material_size_canvas_dictionary = Hash.new
		"A".upto("AL") do |alphabet_character|

			@cell_content = "#{retail_material_size_canvas.cell(1, alphabet_character)}"
			@retail_material_size_canvas_dictionary[@cell_content] = alphabet_character
		end

		# FRAMING, STRETCHING, MATTING
		# Automatically scan the source column names and store them in an associative array
		@retail_framing_stretching_matting_dictionary = Hash.new
		"A".upto("N") do |alphabet_character|

			@cell_content = "#{retail_framing_stretching_matting.cell(2, alphabet_character)}"
			@retail_framing_stretching_matting_dictionary[@cell_content] = alphabet_character
		end


		# Fill every line in the template file up with
		# the right value taken from the source input file		
		@destination_line = 2
		#2.upto(source.last_row) do |source_line|
		4679.upto(4679) do |source_line|


			#{}"A".upto("ER") do |alphabet_character|

				#@source_column = @source_dictionary["#{template.cell(1, alphabet_character)}".downcase.strip]
				#template.set(@destination_line, alphabet_character, @source_column)
				
			#end

		
			# Sku
			@template_column = @template_dictionary["sku"]
			@source_column = @source_dictionary["Item Code"]
			template.set(@destination_line, @template_column, "#{source.cell(source_line, @source_column)}")
			

			@template_column = @template_dictionary["_attribute_set"]
			template.set(@destination_line, @template_column, "Topart - Products")
			

			@template_column = @template_dictionary["_type"]
			template.set(@destination_line, @template_column, "simple")

			@collections_count = 0

			# Artist Focus: look for artists names
			@source_column = @source_dictionary["UDF_ARTIST_NAME"]
			if ( "#{source.cell(source_line,@source_column)}".downcase.strip == "chris donovan" or "#{source.cell(source_line,@source_column)}".downcase.strip == "luke wilson" or "#{source.cell(source_line,@source_column)}".downcase.strip == "erin lange" or "#{source.cell(source_line,@source_column)}".downcase.strip == "gregory williams" or "#{source.cell(source_line,@source_column)}".downcase.strip == "john seba" or "#{source.cell(source_line,@source_column)}".downcase.strip == "mike klung" or "#{source.cell(source_line,@source_column)}".downcase.strip == "alex edwards")

				@template_column = @template_dictionary["_category"]
				template.set(@destination_line + @collections_count, @template_column, "Artist Focus/" + "#{source.cell(source_line,@source_column)}")

				@template_column = @template_dictionary["_root_category"]
				template.set(@destination_line + @collections_count, @template_column, "Root Category")

				@collections_count = @collections_count + 1

			end

			
			# Check if any keyword matches any category name. 
			# If there is a match, add the category name to the corresponding product row.
			2.upto(source_categories.last_row) do |source_categories_line|
			#3.upto(3) do |source_categories_line|

			# Discard the "delete" categories
				if ( "#{source_categories.cell(source_categories_line,'C')}".downcase.strip != "delete" )

					@source_column = @source_dictionary["UDF_ATTRIBUTES"]

					@source_keywords_string = "#{source.cell(source_line, @source_column)}".downcase
					@categories_keywords_array = "#{source_categories.cell(source_categories_line,'B')}".downcase.split(",")
					@category_name = "#{source_categories.cell(source_categories_line,'A')}".strip

					@written_categories = []

					# Browse Art
					0.upto(@categories_keywords_array.size) do |i|

						@string = @categories_keywords_array[i]
						if ( @string )

							@string = @string.strip
							if ( @source_keywords_string.include?(@string) and !@written_categories.include?(@category_name))

								@template_column = @template_dictionary["_category"]
								template.set(@destination_line + @collections_count, @template_column, "Browse Art/" + @category_name)

								@written_categories << @category_name
								
								@template_column = @template_dictionary["_root_category"]
								template.set(@destination_line + @collections_count, @template_column, "Root Category")

								@collections_count = @collections_count + 1

							end
						end
					end
				end
			end


			### Featured Collections ###
			# Floral Patterns
			@source_column = @source_dictionary["UDF_ATTRIBUTES"]
			if "#{source.cell(source_line, @source_column)}".downcase.include? "floral" and "#{source.cell(source_line, @source_column)}".downcase.include? "decorative"

				@template_column = @template_dictionary["_category"]
				template.set(@destination_line + @collections_count, @template_column, "Collections/Featured Collections/Floral Patterns")

				@template_column = @template_dictionary["_root_category"]
				template.set(@destination_line + @collections_count, @template_column, "Root Category")

				@collections_count = @collections_count + 1

			end

			# Contemporary Trends
			@source_column = @source_dictionary["UDF_ATTRIBUTES"]
			if "#{source.cell(source_line, @source_column)}".downcase.include? "contemporary trends"

				@template_column = @template_dictionary["_category"]
				template.set(@destination_line + @collections_count, @template_column, "Collections/Featured Collections/Contemporary Trends")

				@template_column = @template_dictionary["_root_category"]
				template.set(@destination_line + @collections_count, @template_column, "Root Category")

				@collections_count = @collections_count + 1

			end

			# Sandy Escape
			@source_column = @source_dictionary["UDF_ATTRIBUTES"]
			if "#{source.cell(source_line,@source_column)}".downcase.include? "beach"

				@template_column = @template_dictionary["_category"]
				template.set(@destination_line + @collections_count, @template_column, "Collections/Featured Collections/Sandy Escape")

				@template_column = @template_dictionary["_root_category"]
				template.set(@destination_line + @collections_count, @template_column, "Root Category")

				@collections_count = @collections_count + 1

			end

			### End of Featured Collections ###


			# Oversize Variety
			@source_column = @source_dictionary["UDF_OVERSIZE"]
			if ( "#{source.cell(source_line, @source_column)}" == "Y")

				@template_column = @template_dictionary["_category"]
				template.set(@destination_line + @collections_count, @template_column, "Collections/Oversize Variety")

				@template_column = @template_dictionary["_root_category"]
				template.set(@destination_line + @collections_count, @template_column, "Root Category")

				@collections_count = @collections_count + 1

			end

			# Abastract Geometry
			@source_column = @source_dictionary["UDF_ATTRIBUTES"]
			if "#{source.cell(source_line, @source_column)}".downcase.include? "abstract" and "#{source.cell(source_line, @source_column)}".downcase.include? "geometric"

				@template_column = @template_dictionary["_category"]
				template.set(@destination_line + @collections_count, @template_column, "Collections/Abstract Geometry")

				@template_column = @template_dictionary["_root_category"]
				template.set(@destination_line + @collections_count, @template_column, "Root Category")

				@collections_count = @collections_count + 1

			end
			

			# Urban Industrial
			@source_column = @source_dictionary["UDF_ATTRIBUTES"]
			if "#{source.cell(source_line,@source_column)}".downcase.include? "industrial"

				@template_column = @template_dictionary["_category"]
				template.set(@destination_line + @collections_count, @template_column, "Collections/Urban Industrial")

				@template_column = @template_dictionary["_root_category"]
				template.set(@destination_line + @collections_count, @template_column, "Root Category")

				@collections_count = @collections_count + 1

			end

			# Gustav Klimt
			@source_column = @source_dictionary["UDF_ATTRIBUTES"]
			if "#{source.cell(source_line,@source_column)}".downcase.include? "klimt"

				@template_column = @template_dictionary["_category"]
				template.set(@destination_line + @collections_count, @template_column, "Collections/Gustav Klimt-150th Anniversary")

				@template_column = @template_dictionary["_root_category"]
				template.set(@destination_line + @collections_count, @template_column, "Root Category")

				@collections_count = @collections_count + 1

			end

			@template_column = @template_dictionary["_product_websites"]
			template.set(@destination_line, @template_column, "base")
			
			# Alt Size 1, Alt Size 2, Alt Size 3, Alt Size 4
			@source_column = @source_dictionary["UDF_ALTS1"]
			@template_column = @template_dictionary["alt_size_1"]
			template.set(@destination_line, @template_column, "#{source.cell(source_line, @source_column)}")

			@source_column = @source_dictionary["UDF_ALTS2"]
			@template_column = @template_dictionary["alt_size_2"]
			template.set(@destination_line, @template_column, "#{source.cell(source_line, @source_column)}")

			@source_column = @source_dictionary["UDF_ALTS3"]
			@template_column = @template_dictionary["alt_size_3"]
			template.set(@destination_line, @template_column, "#{source.cell(source_line, @source_column)}")

			@source_column = @source_dictionary["UDF_ALTS4"]
			@template_column = @template_dictionary["alt_size_4"]
			template.set(@destination_line, @template_column, "#{source.cell(source_line, @source_column)}")
			
			# Color: Look into "keywords" and search for colors...
			# ...and add each color to the same column but on one @destination_line below
			@source_column = @source_dictionary["UDF_ATTRIBUTES"]
			@template_column = @template_dictionary["color"]
			@color_count = 0;
			0.upto(@color_set.length) do |n|
				if "#{source.cell(source_line, @source_column)}".downcase.include? "#{@color_set[n]}"
					template.set(@destination_line + @color_count, @template_column, "#{@color_set[n]}")
					@color_count = @color_count + 1
				end
			end
			
			
			#Description
			@source_column = @source_dictionary["Description"]
			@template_column = @template_dictionary["description"]
			template.set(@destination_line, @template_column, "#{source.cell(source_line, @source_column)}")


			# Embellishments
			@embellishments_count = 0
			@template_column = @template_dictionary["embellishments"]

			@source_column = @source_dictionary["UDF_METALLICINK"]
			if "#{source.cell(source_line, @source_column)}" == "Y"
				template.set(@destination_line + @embellishments_count, @template_column, "Metallic")
				@embellishments_count = @embellishments_count + 1
			end
			@source_column = @source_dictionary["UDF_FOIL"]
			if "#{source.cell(source_line, @source_column)}" == "Y"
				template.set(@destination_line + @embellishments_count, @template_column, "Foil")
				@embellishments_count = @embellishments_count + 1
			end
			@source_column = @source_dictionary["UDF_SERIGRAPH"]
			if "#{source.cell(source_line, @source_column)}" == "Y"
				template.set(@destination_line + @embellishments_count, @template_column, "Serigraph")
				@embellishments_count = @embellishments_count + 1
			end
			@source_column = @source_dictionary["UDF_EMBOSSED"]
			if "#{source.cell(source_line, @source_column)}" == "Y"
				template.set(@destination_line + @embellishments_count, @template_column, "Embossed")
				@embellishments_count = @embellishments_count + 1
			end


			#Enable Google Checkout
			@template_column = @template_dictionary["enable_googlecheckout"]
			template.set(@destination_line, @template_column, "1")


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
			@source_column = @source_dictionary["UDF_ORIENTATION"]
			@template_column = @template_dictionary["format"]
			template.set(@destination_line, @template_column, "#{source.cell(source_line, @source_column)}")


			#has_options
			@source_column = @source_dictionary["Item Code"]
			@template_column = @template_dictionary["has_options"]
			if "#{source.cell(source_line, @source_column)}" =~ /DG$/ 
				template.set(@destination_line, @template_column, "1")
			else
				template.set(@destination_line, @template_column, "0")
			end

			# Image size cm
			@source_column = @source_dictionary["UDF_IMAGE_SIZE_CM"]
			@template_column = @template_dictionary["image_size_cm"]
			template.set(@destination_line, @template_column, "#{source.cell(source_line, @source_column)}")

			# Image size inches
			@source_column = @source_dictionary["UDF_IMAGE_SIZE_IN"]
			@template_column = @template_dictionary["image_size_inches"]
			template.set(@destination_line, @template_column, "#{source.cell(source_line, @source_column)}")

			#Keywords
			#@source_column = @source_dictionary["UDF_ATTRIBUTES"]
			#@template_column = @template_dictionary["keywords"]
			#@keywords_list = "#{source.cell(source_line, @source_column)}".downcase.split(";").first(20)
			#template.set(@destination_line, @template_column, @keywords_list.join(";"))
			#@keywords_list = "#{source.cell(source_line, @source_column)}".downcase
			#template.set(@destination_line, @template_column, @keywords_list)


			#Meta Description
			@source_column = @source_dictionary["Description"]
			@template_column = @template_dictionary["meta_description"]
			template.set(@destination_line, @template_column, "#{source.cell(source_line, @source_column)}")

			#Meta Kewyords
			@source_column = @source_dictionary["UDF_ATTRIBUTES"]
			@template_column = @template_dictionary["meta_keyword"]
			#@keywords_list = "#{source.cell(source_line, @source_column)}".downcase.split(";").first(20)
			#template.set(@destination_line, @template_column, @keywords_list.join(";"))
			@keywords_list = "#{source.cell(source_line, @source_column)}".downcase
			template.set(@destination_line, @template_column, @keywords_list)

			#Meta title
			@source_column = @source_dictionary["Description"]
			@template_column = @template_dictionary["meta_title"]
			template.set(@destination_line, @template_column, "#{source.cell(source_line, @source_column)}")


			#msrp_display_actual_price_type
			@template_column = @template_dictionary["msrp_display_actual_price_type"]
			template.set(@destination_line, @template_column, "Use config")
			
			#msrp_enabled
			@template_column = @template_dictionary["msrp_enabled"]
			template.set(@destination_line, @template_column, "Use config")

			#Name
			@source_column = @source_dictionary["Description"]
			@template_column = @template_dictionary["name"]
			template.set(@destination_line, @template_column, "#{source.cell(source_line,@source_column)}")

			#options_container
			@template_column = @template_dictionary["options_container"]
			template.set(@destination_line, @template_column, "Block after Info Column")

			#Oversize
			@source_column = @source_dictionary["UDF_OVERSIZE"]
			@template_column = @template_dictionary["over_size"]
			if "#{source.cell(source_line,@source_column)}" == "Y"
				template.set(@destination_line, @template_column, "Yes")
			else
				template.set(@destination_line, @template_column, "No")
			end

			#Paper size cm
			@source_column = @source_dictionary["UDF_PAPER_SIZE_CM"]
			@template_column = @template_dictionary["page_size_cm"]
			template.set(@destination_line, @template_column, "#{source.cell(source_line, @source_column)}")

			#Paper size inches
			@source_column = @source_dictionary["UDF_PAPER_SIZE_IN"]
			@template_column = @template_dictionary["paper_size_inches"]
			template.set(@destination_line, @template_column, "#{source.cell(source_line, @source_column)}")

			#A4POD
			@source_column = @source_dictionary["UDF_A4POD"]
			@template_column = @template_dictionary["pod"]
			if "#{source.cell(source_line,@source_column)}" == "Y"
				template.set(@destination_line, @template_column, "Yes")
			else
				template.set(@destination_line, @template_column, "No")
			end

			#Price
			@source_column = @source_dictionary["SuggestedRetailPrice"]
			@template_column = @template_dictionary["price"]
			template.set(@destination_line, @template_column, "#{source.cell(source_line, @source_column)}")


			#required_options
			@source_column = @source_dictionary["Item Code"]
			@template_column = @template_dictionary["required_options"]
			if "#{source.cell(source_line,@source_column)}" =~ /DG$/ 
				template.set(@destination_line, @template_column, "1")
			else
				template.set(@destination_line, @template_column, "0")
			end

			#Short description
			@source_column = @source_dictionary["Description"]
			@template_column = @template_dictionary["short_description"]
			template.set(@destination_line, @template_column, "#{source.cell(source_line, @source_column)}")


			#Size category: for posters
			@source_column = @source_dictionary["UDF_IMAGE_SIZE_CM"]
			@template_column = @template_dictionary["size_category"]
			@image_size_cm = "#{source.cell(source_line,@source_column)}"
			
			@width = @image_size_cm.gsub(/ x .[0-9]/, "")
			@height = @image_size_cm.gsub(/.[0-9] x /, "")

			#Convert UI to inches to have a consistent comparison with the spreadsheet
			@ui = ( (@width.to_i + @height.to_i) / 2.54).to_i;

			if (@ui != 0)

				if @ui < 40 
					template.set(@destination_line, @template_column, "Petite")
				end

				if @ui >= 40 and @ui <  50
					template.set(@destination_line, @template_column, "Small")
				end

				if @ui >= 50 and @ui < 60 
					template.set(@destination_line, @template_column, "Medium")
				end

				if @ui >= 60 and @ui < 70
					template.set(@destination_line, @template_column, "Large")
				end

				if @ui >= 70   
					template.set(@destination_line, @template_column, "Oversize")
				end

			end

			#Status: enabled (1), disabled (2)
			@source_column = @source_dictionary["Item Code"]
			@template_column = @template_dictionary["status"]
			if "#{source.cell(source_line,@source_column)}" =~ /DG$/ 
				template.set(@destination_line, @template_column, "1")
			else
				template.set(@destination_line, @template_column, "1")
			end

			#Tax class ID
			@template_column = @template_dictionary["tax_class_id"]
			template.set(@destination_line, @template_column, "2")

			#total_quantity_on_hand
			@source_column = @source_dictionary["Item Code"]
			@template_column = @template_dictionary["total_quantity_on_hand"]
			if "#{source.cell(source_line,@source_column)}" =~ /DG$/ 
				template.set(@destination_line, @template_column, "0".to_i)
			else
				@source_column_2 = @source_dictionary["TotalQuantityOnHand"]
				template.set(@destination_line, @template_column, "#{source.cell(source_line,@source_column_2)}".to_i)
			end

			#udf_anycustom
			@source_column = @source_dictionary["udf_anycustom"]
			@template_column = @template_dictionary["udf_anycustom"]
			if "#{source.cell(source_line,@source_column)}" == "Y"
				template.set(@destination_line, @template_column, "Yes")
			else
				template.set(@destination_line, @template_column, "No")
			end

			#Artist name
			@source_column = @source_dictionary["UDF_ARTIST_NAME"]
			@template_column = @template_dictionary["udf_artist_name"]
			template.set(@destination_line, @template_column, "#{source.cell(source_line, @source_column)}")

			#Copyright
			@source_column = @source_dictionary["UDF_COPYRIGHT"]
			@template_column = @template_dictionary["udf_copyright"]
			if "#{source.cell(source_line, @source_column)}" == "Y"
				template.set(@destination_line, @template_column, "Yes")
			else
				template.set(@destination_line, @template_column, "No")
			end

			#udf_crimage
			@source_column = @source_dictionary["UDF_CRIMAGE"]
			@template_column = @template_dictionary["udf_crimage"]
			template.set(@destination_line, @template_column, "#{source.cell(source_line, @source_column)}")

			#udf_crline
			@source_column = @source_dictionary["UDF_CRLINE"]
			@template_column = @template_dictionary["udf_crline"]
			template.set(@destination_line, @template_column, "#{source.cell(source_line, @source_column)}")

			#udf_dnd
			@source_column = @source_dictionary["UDF_DND"]
			@template_column = @template_dictionary["udf_dnd"]
			if "#{source.cell(source_line, @source_column)}" == "Y"
				template.set(@destination_line, @template_column, "Yes")
			else
				template.set(@destination_line, @template_column, "No")
			end

			#udf_embellished
			@source_column = @source_dictionary["UDF_EMBELLISHED"]
			@template_column = @template_dictionary["udf_embellished"]
			if "#{source.cell(source_line,@source_column)}" == "Y"
				template.set(@destination_line, @template_column, "Yes")
			else
				template.set(@destination_line, @template_column, "No")
			end

			#udf_framed
			@source_column = @source_dictionary["UDF_FRAMED"]
			@template_column = @template_dictionary["udf_framed"]
			if "#{source.cell(source_line,@source_column)}" == "Y"
				template.set(@destination_line, @template_column, "Yes")
			else
				template.set(@destination_line, @template_column, "No")
			end

			#udf_imsource
			@source_column = @source_dictionary["UDF_IMSOURCE"]
			@template_column = @template_dictionary["udf_imsource"]
			template.set(@destination_line, @template_column, "#{source.cell(source_line,@source_column)}")

			#udf_limited
			@source_column = @source_dictionary["UDF_LIMITED"]
			@template_column = @template_dictionary["udf_limited"]
			if "#{source.cell(source_line,@source_column)}" == "Y"
				template.set(@destination_line, @template_column, "Yes")
			else
				template.set(@destination_line, @template_column, "No")
			end

			#udf_new
			@source_column = @source_dictionary["UDF_NEW"]
			@template_column = @template_dictionary["udf_new"]
			if "#{source.cell(source_line,@source_column)}" == "Y"
				template.set(@destination_line, @template_column, "Yes")
			else
				template.set(@destination_line, @template_column, "No")
			end

			#udf_osdp
			@source_column = @source_dictionary["UDF_OSDP"]
			@template_column = @template_dictionary["udf_osdp"]
			if "#{source.cell(source_line,@source_column)}" == "Y"
				template.set(@destination_line, @template_column, "Yes")
			else
				template.set(@destination_line, @template_column, "No")
			end

			#udf_pricecode
			@source_column = @source_dictionary["UDF_PRICECODE"]
			@template_column = @template_dictionary["udf_pricecode"]
			template.set(@destination_line, @template_column, "#{source.cell(source_line,@source_column)}")

			#udf_ratiocode
			@source_column = @source_dictionary["UDF_RATIOCODE"]
			@template_column = @template_dictionary["udf_ratiocode"]
			time = "#{source.cell(source_line,@source_column)}".to_i
			hours = time/3600
			minutes = (time/60 - hours * 60)

			ratio_code = hours.to_s << ":" << minutes.to_s
			template.set(@destination_line, @template_column, ratio_code)

			#udf_tar
			@source_column = @source_dictionary["UDF_TAR"]
			@template_column = @template_dictionary["udf_tar"]
			if "#{source.cell(source_line,@source_column)}" == "Y"
				template.set(@destination_line, @template_column, "Yes")
			else
				template.set(@destination_line, @template_column, "No")
			end

			#URL Key, with the SKU as suffix to keep it unique among products
			@source_column_1 = @source_dictionary["Description"]
			@source_column_2 = @source_dictionary["Item Code"]
			@template_column = @template_dictionary["url_key"]
			template.set(@destination_line, @template_column, "#{source.cell(source_line,@source_column_1)}".gsub(/[ ]/, '-')  << "-" << "#{source.cell(source_line,@source_column_2)}")

			#Visibility
			@template_column = @template_dictionary["visibility"]
			template.set(@destination_line, @template_column, "4")

			#Weight
			@template_column = @template_dictionary["weight"]
			template.set(@destination_line, @template_column, "1")

			#Qty
			@source_column_1 = @source_dictionary["Item Code"]
			@source_column_2 = @source_dictionary["TotalQuantityOnHand"]
			@template_column = @template_dictionary["qty"]
			if "#{source.cell(source_line,@source_column_1)}" =~ /DG$/ 
				template.set(@destination_line, @template_column, "0")
			else
				template.set(@destination_line, @template_column, "#{source.cell(source_line,@source_column_2)}")
			end

			#Min qty
			@template_column = @template_dictionary["min_qty"]
			template.set(@destination_line, @template_column, "0")

			#use_config_min_qty
			@template_column = @template_dictionary["use_config_min_qty"]
			template.set(@destination_line, @template_column, "1")

			#is_qty_decimal
			@template_column = @template_dictionary["is_qty_decimal"]
			template.set(@destination_line, @template_column, "0")

			#backorders
			@template_column = @template_dictionary["backorders"]
			template.set(@destination_line, @template_column, "0")

			#use_config_backorders
			@template_column = @template_dictionary["use_config_backorders"]
			template.set(@destination_line, @template_column, "1")

			#min_sale_qty
			@template_column = @template_dictionary["min_sale_qty"]
			template.set(@destination_line, @template_column, "1")

			#use_config_min_sale_qty
			@template_column = @template_dictionary["use_config_min_sale_qty"]
			template.set(@destination_line, @template_column, "1")

			#max_sale_qty
			@template_column = @template_dictionary["max_sale_qty"]
			template.set(@destination_line, @template_column, "0")

			#use_config_max_sale_qty
			@template_column = @template_dictionary["use_config_max_sale_qty"]
			template.set(@destination_line, @template_column, "1")

			#is_in_stock
			@source_column = @source_dictionary["Item Code"]
			@template_column = @template_dictionary["is_in_stock"]
			if "#{source.cell(source_line,@source_column)}" =~ /DG$/
				template.set(@destination_line, @template_column, "1")
			else
				template.set(@destination_line, @template_column, "0")
			end

			#use_config_notify_stock_qty
			@source_column = @source_dictionary["Item Code"]
			@template_column = @template_dictionary["use_config_notify_stock_qty"]
			if "#{source.cell(source_line,@source_column)}" =~ /DG$/
				template.set(@destination_line, @template_column, "1")
			else
				template.set(@destination_line, @template_column, "0")
			end

			#manage_stock
			@source_column = @source_dictionary["Item Code"]
			@template_column = @template_dictionary["manage_stock"]
			if "#{source.cell(source_line,@source_column)}" =~ /DG$/
				template.set(@destination_line, @template_column, "0")
			else
				template.set(@destination_line, @template_column, "1")
			end

			#use_config_manage_stock
			@source_column = @source_dictionary["Item Code"]
			@template_column = @template_dictionary["use_config_manage_stock"]
			if "#{source.cell(source_line,@source_column)}" =~ /DG$/
				template.set(@destination_line, @template_column, "0")
			else
				template.set(@destination_line, @template_column, "1")
			end

			#stock_status_changed_auto
			@template_column = @template_dictionary["stock_status_changed_auto"]
			template.set(@destination_line, @template_column, "0")

			#use_config_qty_increments
			@template_column = @template_dictionary["use_config_qty_increments"]
			template.set(@destination_line, @template_column, "1")

			#qty_increments
			@template_column = @template_dictionary["qty_increments"]
			template.set(@destination_line, @template_column, "0")

			#use_config_enable_qty_inc
			@template_column = @template_dictionary["use_config_enable_qty_inc"]
			template.set(@destination_line, @template_column, "1")

			#enable_qty_increments
			@template_column = @template_dictionary["enable_qty_increments"]
			template.set(@destination_line, @template_column, "0")

			#is_decimal_divided
			@template_column = @template_dictionary["is_decimal_divided"]
			template.set(@destination_line, @template_column, "0")



			########## Custom options columns ##########

			@source_column = @source_dictionary["Item Code"]
			if "#{source.cell(source_line,@source_column)}" =~ /DG$/

				# MATERIAL: paper and canvas are static hard-coded options.

				########### Material ###############

				#_custom_option_type
				@template_column = @template_dictionary["_custom_option_type"]
				template.set(@destination_line, @template_column, "radio")
				#_custom_option_title
				@template_column = @template_dictionary["_custom_option_title"]
				template.set(@destination_line, @template_column, "Material")
				#_custom_option_is_required
				@template_column = @template_dictionary["_custom_option_is_required"]
				template.set(@destination_line, @template_column, "1")
				#_custom_option_max_characters
				@template_column = @template_dictionary["_custom_option_max_characters"]
				template.set(@destination_line, @template_column, "0")
				#_custom_option_sort_order
				@template_column = @template_dictionary["_custom_option_sort_order"]
				template.set(@destination_line, @template_column, "0")
				

				#_custom_option_row_title
				@template_column = @template_dictionary["_custom_option_row_title"]
				template.set(@destination_line, @template_column, "Paper")
				#_custom_option_row_price
				@template_column = @template_dictionary["_custom_option_row_price"]
				template.set(@destination_line, @template_column, "0.00")
				#_custom_option_row_sku
				@template_column = @template_dictionary["_custom_option_row_sku"]
				template.set(@destination_line, @template_column, "material_paper")
				#_custom_option_row_sort
				@template_column = @template_dictionary["_custom_option_row_sort"]
				template.set(@destination_line, @template_column, "0")

				@destination_line = @destination_line + 1

				#_custom_option_row_title
				@template_column = @template_dictionary["_custom_option_row_title"]
				template.set(@destination_line, @template_column, "Canvas")
				#_custom_option_row_price
				@template_column = @template_dictionary["_custom_option_row_price"]
				template.set(@destination_line, @template_column, "0.00")
				#_custom_option_row_sku
				@template_column = @template_dictionary["_custom_option_row_sku"]
				template.set(@destination_line, @template_column, "material_canvas")
				#_custom_option_row_sort
				@template_column = @template_dictionary["_custom_option_row_sort"]
				template.set(@destination_line, @template_column, "1")

				@destination_line = @destination_line + 1

				########### End of Material ###############


				#############SIZE#############

				#_custom_option_type
				@template_column = @template_dictionary["_custom_option_type"]
				template.set(@destination_line, @template_column, "radio")
				#_custom_option_title
				@template_column = @template_dictionary["_custom_option_title"]
				template.set(@destination_line, @template_column, "Size")
				#_custom_option_is_required
				@template_column = @template_dictionary["_custom_option_is_required"]
				template.set(@destination_line, @template_column, "1")
				#_custom_option_max_characters
				@template_column = @template_dictionary["_custom_option_max_characters"]
				template.set(@destination_line, @template_column, "0")
				#_custom_option_sort_order
				@template_column = @template_dictionary["_custom_option_sort_order"]
				template.set(@destination_line, @template_column, "1")
				
				# We need to extract the right prices, looking them up by (i.e. matching) the ratio column

				# Extract and map the border treatments:
				# 1) Scan for every row into the master paper and master canvas sheets
				# 2) check if the ratio matches the one contained in the product attribute 
				# 3) If the 2 ratios match, then copy the specific retail price option
				
				@source_column = @source_dictionary["UDF_RATIOCODE"]
				time = "#{source.cell(source_line, @source_column)}".to_i
				hours = time/3600
				minutes = (time/60 - hours * 60)

				@source_ratio_code = hours.to_s << ":" << minutes.to_s

				@match_index = 0

				# Master Paper Sheet
				2.upto(retail_material_size_paper.last_row) do |retail_line|

					@retail_column = @retail_material_size_paper_dictionary["Ratio"]
					@retail_ratio_code = "#{retail_material_size_paper.cell(retail_line, @retail_column)}"


					# Check for available sizes
					if @source_ratio_code == @retail_ratio_code

						#p "source ratio code: " + @source_ratio_code
						#p "retail ratio code: " + @retail_ratio_code
						#p "______________________"
						
						@retail_column = @retail_material_size_paper_dictionary["Ratio"]
						@retail_ratio_code = "#{retail_material_size_paper.cell(retail_line, @retail_column)}"

						@retail_column = @retail_material_size_paper_dictionary["SIZE DESCRIPTION"]
						@size_name = "#{retail_material_size_paper.cell(retail_line, @retail_column)}"

						@retail_column = @retail_material_size_paper_dictionary["Rolled Photo Paper Retail"]
						@size_price = "#{retail_material_size_paper.cell(retail_line, @retail_column)}"

						@retail_column = @retail_material_size_paper_dictionary["UI"]
						@size_paper_ui = "#{retail_material_size_paper.cell(retail_line, @retail_column)}".to_i


						#_custom_option_row_title
						@template_column = @template_dictionary["_custom_option_row_title"]
						template.set(@destination_line, @template_column, @size_name)
						#_custom_option_row_price
						@template_column = @template_dictionary["_custom_option_row_price"]
						template.set(@destination_line, @template_column, @size_price)
						#_custom_option_row_sku
						@template_column = @template_dictionary["_custom_option_row_sku"]
						template.set(@destination_line, @template_column, "size_paper_" + @size_name.downcase + "_ui_" + @size_paper_ui.to_s)
						#_custom_option_row_sort
						@template_column = @template_dictionary["_custom_option_row_sort"]
						template.set(@destination_line, @template_column, @match_index)

						@destination_line = @destination_line + 1

						@match_index = @match_index + 1

					end

				end

				# Master Canvas Sheet
				2.upto(retail_material_size_canvas.last_row) do |retail_line|

					@retail_column = @retail_material_size_canvas_dictionary["Ratio"]
					@retail_ratio_code = "#{retail_material_size_canvas.cell(retail_line, @retail_column)}"
						
						#p "source ratio code: " + @source_ratio_code
						#p "retail ratio code: " + @retail_ratio_code
						#p "______________________"
					
					@count = 0

					# Check for available sizes and border treatments prices
					if @source_ratio_code == @retail_ratio_code

						@retail_column = @retail_material_size_canvas_dictionary["SIZE DESCRIPTION"]
						@size_name = "#{retail_material_size_canvas.cell(retail_line, @retail_column)}"	
						
						@retail_column = @retail_material_size_canvas_dictionary["RETAIL PRICE"]
						@size_price_treatment_1 = "#{retail_material_size_canvas.cell(retail_line, @retail_column)}"
						
						@retail_column = @retail_material_size_canvas_dictionary['ROLLED CANVAS 2" BLACK Border RETAIL']
						@size_price_treatment_2 = "#{retail_material_size_canvas.cell(retail_line, @retail_column)}"

						@retail_column = @retail_material_size_canvas_dictionary['ROLLED CANVAS 2" MIRROR Border RETAIL']
						@size_price_treatment_3 = "#{retail_material_size_canvas.cell(retail_line, @retail_column)}"
						
						#p @size_price_treatment_1
						#p @size_price_treatment_2
						#p @size_price_treatment_3
						
						@size_prices = Array.new
						@size_prices << @size_price_treatment_1 << @size_price_treatment_2 << @size_price_treatment_3

						@retail_column = @retail_material_size_canvas_dictionary["UI"]
						@size_canvas_ui = "#{retail_material_size_canvas.cell(retail_line, @retail_column)}".to_i

						0.upto(2) do |count|

							#_custom_option_row_title
							@template_column = @template_dictionary["_custom_option_row_title"]
							template.set(@destination_line, @template_column, @size_name)
							#_custom_option_row_price
							@template_column = @template_dictionary["_custom_option_row_price"]
							template.set(@destination_line, @template_column, @size_prices[count])
							#_custom_option_row_sku
							@template_column = @template_dictionary["_custom_option_row_sku"]
							template.set(@destination_line, @template_column, "size_canvas_" + @size_name.downcase + "_treatment_" + (count+1).to_s + "_ui_" + @size_canvas_ui.to_s)
							#_custom_option_row_sort
							@template_column = @template_dictionary["_custom_option_row_sort"]
							template.set(@destination_line, @template_column, @match_index + count)

							@destination_line = @destination_line + 1

							@count = count
						
						end

						@match_index = @match_index + 1 + @count

					end

				end



				########### Border Treatments ###############
				# Border Treatments and Stretching options (including names) are static

				#_custom_option_type
				@template_column = @template_dictionary["_custom_option_type"]
				template.set(@destination_line, @template_column, "radio")
				#_custom_option_title
				@template_column = @template_dictionary["_custom_option_title"]
				template.set(@destination_line, @template_column, "Treatments")
				#_custom_option_is_required
				@template_column = @template_dictionary["_custom_option_is_required"]
				template.set(@destination_line, @template_column, "1")
				#_custom_option_max_characters
				@template_column = @template_dictionary["_custom_option_max_characters"]
				template.set(@destination_line, @template_column, "0")
				#_custom_option_sort_order
				@template_column = @template_dictionary["_custom_option_sort_order"]
				template.set(@destination_line, @template_column, "1")

				#_custom_option_row_title
				@template_column = @template_dictionary["_custom_option_row_title"]
				template.set(@destination_line, @template_column, "None")
				#_custom_option_row_price
				@template_column = @template_dictionary["_custom_option_row_price"]
				template.set(@destination_line, @template_column, "0")
				#_custom_option_row_sku
				@template_column = @template_dictionary["_custom_option_row_sku"]
				template.set(@destination_line, @template_column, "treatments_none")
				#_custom_option_row_sort
				@template_column = @template_dictionary["_custom_option_row_sort"]
				template.set(@destination_line, @template_column, "0")

				@destination_line = @destination_line + 1
				

				#_custom_option_row_title
				@template_column = @template_dictionary["_custom_option_row_title"]
				template.set(@destination_line, @template_column, "3\" white")
				#_custom_option_row_price
				@template_column = @template_dictionary["_custom_option_row_price"]
				template.set(@destination_line, @template_column, "0")
				#_custom_option_row_sku
				@template_column = @template_dictionary["_custom_option_row_sku"]
				template.set(@destination_line, @template_column, "border_treatment_3_inches_of_white")
				#_custom_option_row_sort
				@template_column = @template_dictionary["_custom_option_row_sort"]
				template.set(@destination_line, @template_column, "1")

				@destination_line = @destination_line + 1

				#_custom_option_row_title
				@template_column = @template_dictionary["_custom_option_row_title"]
				template.set(@destination_line, @template_column, "2\" black 1\" white")
				#_custom_option_row_price
				@template_column = @template_dictionary["_custom_option_row_price"]
				template.set(@destination_line, @template_column, "0")
				#_custom_option_row_sku
				@template_column = @template_dictionary["_custom_option_row_sku"]
				template.set(@destination_line, @template_column, "border_treatment_2_inches_of_black_and_1_inch_of_white")
				#_custom_option_row_sort
				@template_column = @template_dictionary["_custom_option_row_sort"]
				template.set(@destination_line, @template_column, "2")

				@destination_line = @destination_line + 1

				#_custom_option_row_title
				@template_column = @template_dictionary["_custom_option_row_title"]
				template.set(@destination_line, @template_column, "2\" mirrored 1\" white")
				#_custom_option_row_price
				@template_column = @template_dictionary["_custom_option_row_price"]
				template.set(@destination_line, @template_column, "0")
				#_custom_option_row_sku
				@template_column = @template_dictionary["_custom_option_row_sku"]
				template.set(@destination_line, @template_column, "border_treatment_2_inches_mirrored_and_1_inch_of_white")
				#_custom_option_row_sort
				@template_column = @template_dictionary["_custom_option_row_sort"]
				template.set(@destination_line, @template_column, "3")

				@destination_line = @destination_line + 1
				


				########### FRAMING ###########
				# Master Framing

				#_custom_option_type
				@template_column = @template_dictionary["_custom_option_type"]
				template.set(@destination_line, @template_column, "drop_down")
				#_custom_option_title
				@template_column = @template_dictionary["_custom_option_title"]
				template.set(@destination_line, @template_column, "Frame")
				#_custom_option_is_required
				@template_column = @template_dictionary["_custom_option_is_required"]
				template.set(@destination_line, @template_column, "1")
				#_custom_option_max_characters
				@template_column = @template_dictionary["_custom_option_max_characters"]
				template.set(@destination_line, @template_column, "0")
				#_custom_option_sort_order
				@template_column = @template_dictionary["_custom_option_sort_order"]
				template.set(@destination_line, @template_column, "4")

				@frame_count = 0;

				2.upto(retail_framing_stretching_matting.last_row) do |retail_line|


					@retail_column = @retail_framing_stretching_matting_dictionary["Descripton"]
					@frame_name = "#{retail_framing_stretching_matting.cell(retail_line, @retail_column)}"

					@frame_name = @frame_name.downcase.tr(" ", "_")

					@retail_column = @retail_framing_stretching_matting_dictionary["United Inch TAR Retail"]
					@frame_ui_price = "#{retail_framing_stretching_matting.cell(retail_line, @retail_column)}"

					@retail_column = @retail_framing_stretching_matting_dictionary["Flat Mounting Cost"]
					@frame_flat_mounting_price = "#{retail_framing_stretching_matting.cell(retail_line, @retail_column)}"

					########### Canvas Stretching ###############
					if @frame_name == "1.5\"_stretcher_bars"
		
						#_custom_option_type
						@template_column = @template_dictionary["_custom_option_type"]
						template.set(@destination_line, @template_column, "checkbox")
						#_custom_option_title
						@template_column = @template_dictionary["_custom_option_title"]
						template.set(@destination_line, @template_column, "Canvas Stretching")
						#_custom_option_is_required
						@template_column = @template_dictionary["_custom_option_is_required"]
						template.set(@destination_line, @template_column, "0")
						#_custom_option_max_characters
						@template_column = @template_dictionary["_custom_option_max_characters"]
						template.set(@destination_line, @template_column, "0")
						#_custom_option_sort_order
						@template_column = @template_dictionary["_custom_option_sort_order"]
						template.set(@destination_line, @template_column, "3")
						

						#_custom_option_row_title
						@template_column = @template_dictionary["_custom_option_row_title"]
						template.set(@destination_line, @template_column, "Stretch it for me")
						#_custom_option_row_price
						@template_column = @template_dictionary["_custom_option_row_price"]
						template.set(@destination_line, @template_column, @frame_ui_price)
						#_custom_option_row_sku
						@template_column = @template_dictionary["_custom_option_row_sku"]
						template.set(@destination_line, @template_column, "canvas_stretching")
						#_custom_option_row_sort
						@template_column = @template_dictionary["_custom_option_row_sort"]
						template.set(@destination_line, @template_column, "0")

						@destination_line = @destination_line + 1

					end


					@retail_column = @retail_framing_stretching_matting_dictionary["Available for Paper"]
					@frame_for_paper = "#{retail_framing_stretching_matting.cell(retail_line, @retail_column)}"

					@retail_column = @retail_framing_stretching_matting_dictionary["Available for Canvas"]
					@frame_for_canvas = "#{retail_framing_stretching_matting.cell(retail_line, @retail_column)}"

					

					# FRAMING: check if the description contains the substring "Frame"
					if @frame_name.downcase.include?("frame")

						#_custom_option_row_title
						@template_column = @template_dictionary["_custom_option_row_title"]
						template.set(@destination_line, @template_column, @frame_name)
						#_custom_option_row_price
						@template_column = @template_dictionary["_custom_option_row_price"]
						template.set(@destination_line, @template_column, @frame_ui_price)
						
						# Available for Paper
						if @frame_for_paper.downcase == "y"
							#_custom_option_row_sku
							@template_column = @template_dictionary["_custom_option_row_sku"]
							template.set(@destination_line, @template_column, "frame_paper" + @frame_name.downcase)
						end

						# Available for Canvas
						if @frame_for_canvas.downcase == "y"
							#_custom_option_row_sku
							@template_column = @template_dictionary["_custom_option_row_sku"]
							template.set(@destination_line, @template_column, "frame_canvas" + @frame_name.downcase)
						end

						#_custom_option_row_sort
						@template_column = @template_dictionary["_custom_option_row_sort"]
						template.set(@destination_line, @template_column, @frame_count)

						@destination_line = @destination_line + 1

						@frame_count = @frame_count + 1

					end

				end


			end	
			
			
			
			# Compute the maximum count among all the multi select options
			# then add it to the destination line count for the next product to be written
			
			@custom_options_array_size = 0

			@multi_select_options = Array.new
			@multi_select_options << @color_count << @embellishments_count << @collections_count

			@source_column = @source_dictionary["Item Code"]
			if "#{source.cell(source_line,@source_column)}" =~ /DG$/
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
