require 'rubygems'
require 'roo'
require 'debugger'

class TemplatesController < ApplicationController

	# GET /generate_template
	# GET /generate_template.json
	def index

 
		# Load the source Excel file, with all the special products info
		#source = Excel.new("http://beta.topart.com/csv/Template_2012_11_28/source.xls")
		source = Excel.new("Template_2013_05_10/source.xls")
		source.default_sheet = source.sheets.first
		
		# Load the Magento template, which is in Open Office format
		#template = Openoffice.new("http://beta.topart.com/csv/Template_2012_11_28/template.ods")
		template = Openoffice.new("Template_2013_05_10/template.ods")
		template.default_sheet = template.sheets.first

		# Automatically scan the template column names and store them in an associative array
		@template_dictionary = Hash.new
		"A".upto("GC") do |alphabet_character|
			@cell_content = "#{template.cell(1, alphabet_character)}"
			@template_dictionary[@cell_content] = alphabet_character
		end

		p "Template headers loaded."

		# Automatically scan the source column names and store them in an associative array
		@source_dictionary = Hash.new
		"A".upto("BU") do |alphabet_character|
			@cell_content = "#{source.cell(1, alphabet_character)}"
			@source_dictionary[@cell_content] = alphabet_character
		end

		p "Source headers loaded."

		# Load the retail_material_size spreadsheet file for paper
		retail_photo_paper = Excel.new("Template_2013_05_10/retail_master.xls")
		retail_photo_paper.default_sheet = retail_photo_paper.sheets[0]

		# Load the retail_material_size spreadsheet file for canvas
		retail_canvas = Excel.new("Template_2013_05_10/retail_master.xls")
		retail_canvas.default_sheet = retail_canvas.sheets[2]

		# Load the retail_framing spreadsheet file to extract framing, stretching and matting information
		retail_framing = Excel.new("Template_2013_05_10/retail_master.xls")
		retail_framing.default_sheet = retail_framing.sheets[3]


		# MATERIAL -> PAPER
		# Automatically scan the source column names and store them in an associative array
		@retail_photo_paper_dictionary = Hash.new
		"A".upto("T") do |alphabet_character|
			@cell_content = "#{retail_photo_paper.cell(1, alphabet_character)}"
			@retail_photo_paper_dictionary[@cell_content] = alphabet_character
		end

		p "Retail photo paper headers correctly loaded."

		# MATERIAL -> CANVAS
		# Automatically scan the source column names and store them in an associative array
		@retail_canvas_dictionary = Hash.new
		"A".upto("AO") do |alphabet_character|
			@cell_content = "#{retail_canvas.cell(1, alphabet_character)}"
			@retail_canvas_dictionary[@cell_content] = alphabet_character
		end

		p "Retail canvas headers correctly loaded."

		@retail_dictionary = Hash.new
		"A".upto("R") do |alphabet_character|
			@cell_content = "#{retail_framing.cell(1, alphabet_character)}"
			@retail_dictionary[@cell_content] = alphabet_character
		end

		
		# FRAMING, STRETCHING, MATTING
		# Automatically scan the source column names and store them in an associative array
		# Declare and fill the retail framing table
		@retail_framing_table = Array.new(retail_framing.last_row, 18)
		i = 0

		# Scan all the source rows and process the F21066 items only, and only once at the beginning for efficiency
		2.upto(source.last_row) do |source_line|

			@primary_vendor_no = "#{source.cell(source_line, @source_dictionary["PrimaryVendorNo"])}"

			if @primary_vendor_no == "F21066"
				@retail_framing_table[i] = Hash.new

				# Store all the MAS specific fields, which means the majority of them
				"A".upto("R") do |alphabet_character|
					@header = "#{retail_framing.cell(1, alphabet_character)}"
					@retail_framing_table[i][@header] = "#{source.cell(source_line, @source_dictionary[@header])}"
				end

				# Store the spreadsheet retail prices only
				2.upto(retail_framing.last_row) do |k|
					#@retail_framing_table[i] = Hash.new

					"C".upto("F") do |alphabet_character|
						@cell_content = "#{retail_framing.cell(1, alphabet_character)}"

						if @retail_framing_table[i]["Item Code"] == "#{retail_framing.cell(k, @retail_dictionary["Item Code"])}"
							@retail_framing_table[i][@cell_content] = "#{retail_framing.cell(k, alphabet_character)}"
						end
					end
				end

				i = i + 1

			end

		end

		p "The F21066 items have been correctly loaded."


		# Load a hash table with all the item codes from the products spreadsheet. Used to check the presence of DGs and corresponding posters
		@item_source_line = Hash.new

		2.upto(10) do |source_line|
		#2.upto(source.last_row) do |source_line|
			@item_code = "#{source.cell(source_line, @source_dictionary["Item Code"])}"
			@item_source_line[@item_code] = source_line
		end

		# We use the following hash table to track DG products that should contain the additional poster size as a custom option
		@posters_and_dgs_hash_table = Hash.new
		@poster_only_hash_table = Hash.new

		@destination_line = 2
		@template_counter = 1

		source_line = 2

		#while source_line <= source.last_row
		#while source_line <= 10
		9570.upto(9708) do |source_line|
				
			### Fields variables for each product are all assigned here ###

			@udf_tar = "#{source.cell(source_line, @source_dictionary["UDF_TAR"])}"

			# Skip importing items where udf_tar = N
			if @udf_tar == "N"
				next
			end

			@primary_vendor_no = "#{source.cell(source_line, @source_dictionary["PrimaryVendorNo"])}"




			if @primary_vendor_no == "F21066"

				template.set(@destination_line, @template_dictionary["primary_vendor_no"], "#{source.cell(source_line, @source_dictionary["PrimaryVendorNo"])}")
				template.set(@destination_line, @template_dictionary["sku"], "#{source.cell(source_line, @source_dictionary["Item Code"])}")
				template.set(@destination_line, @template_dictionary["description"], "#{source.cell(source_line, @source_dictionary["Description"])}")
				template.set(@destination_line, @template_dictionary["udf_fmaxsssin"], "#{source.cell(source_line, @source_dictionary["UDF_FMAXSSSIN"])}")
				template.set(@destination_line, @template_dictionary["udf_fmaxslsin"], "#{source.cell(source_line, @source_dictionary["UDF_FMAXSLSIN"])}")
				template.set(@destination_line, @template_dictionary["udf_fmaxssxcm"], "#{source.cell(source_line, @source_dictionary["UDF_FMAXSSXCM"])}")
				template.set(@destination_line, @template_dictionary["udf_fmaxslscm"], "#{source.cell(source_line, @source_dictionary["UDF_FMAXSLSCM"])}")


				if "#{source.cell(source_line, @source_dictionary["UDF_ECO"])}" == "Y"
					template.set(@destination_line, @template_dictionary["udf_eco"], "Yes")
				else
					template.set(@destination_line, @template_dictionary["udf_eco"], "No")
				end
				
				if "#{source.cell(source_line, @source_dictionary["UDF_FMAVAIL4PAPER"])}" == "Y"
					template.set(@destination_line, @template_dictionary["udf_f_m_avail_4_paper"], "Yes")
				else
					template.set(@destination_line, @template_dictionary["udf_f_m_avail_4_paper"], "No")
				end

				if "#{source.cell(source_line, @source_dictionary["UDF_FMAVAIL4CANVAS"])}" == "Y"
					template.set(@destination_line, @template_dictionary["udf_f_m_avail_4_canvas"], "Yes")
				else
					template.set(@destination_line, @template_dictionary["udf_f_m_avail_4_canvas"], "No")
				end

				
				

				template.set(@destination_line, @template_dictionary["udf_framecat"], "#{source.cell(source_line, @source_dictionary["UDF_FRAMECAT"])}")
				template.set(@destination_line, @template_dictionary["udf_colorcode"], "#{source.cell(source_line, @source_dictionary["UDF_COLORCODE"])}")
				template.set(@destination_line, @template_dictionary["udf_moulding_width"], "#{source.cell(source_line, @source_dictionary["UDF_MOULDINGWIDTH"])}")
				template.set(@destination_line, @template_dictionary["udf_entitytype"], "#{source.cell(source_line, @source_dictionary["UDF_ENTITYTYPE"])}")

				template.set(@destination_line, @template_dictionary["_attribute_set"], "Topart - Products")
				template.set(@destination_line, @template_dictionary["_type"], "simple")
				template.set(@destination_line, @template_dictionary["name"], "#{source.cell(source_line, @source_dictionary["Description"])}")
				template.set(@destination_line, @template_dictionary["price"], "0.0")
				template.set(@destination_line, @template_dictionary["short_description"], "#{source.cell(source_line, @source_dictionary["Description"])}")
				template.set(@destination_line, @template_dictionary["visibility"], "1")
				template.set(@destination_line, @template_dictionary["weight"], "1")
				template.set(@destination_line, @template_dictionary["tax_class_id"], "2")
				template.set(@destination_line, @template_dictionary["status"], "1")

				@destination_line = @destination_line + 1

			end

		end

		p "Creating Framing Template..."
		template.to_csv("framing_inventory.csv")

	end
end

			