require 'rubygems'
require 'roo'
require 'ruby-prof'

class TemplatesController < ApplicationController


	# Example image size = "23 5/8 x 47 1/4"
	def compute_image_size_width(original_size_value)

		image_size_width = 0

		original_image_size_width = original_size_value.slice(0..original_size_value.index("x") - 1)

		if original_image_size_width.include?('/')

			width_fraction_numerator = original_image_size_width.slice(original_image_size_width.index("/") - 1).to_f
			width_fraction_denominator = original_image_size_width.slice(original_image_size_width.index("/") + 1).to_f
			
			image_size_width += width_fraction_numerator/width_fraction_denominator

		end

		if original_image_size_width.include?('.')
			image_size_width += original_image_size_width[0,5].to_f
		else
			image_size_width += original_image_size_width[0,2].to_f
		end

		return image_size_width

	end

	# Example image size = "23 5/8 x 47 1/4"
	def compute_image_size_length(original_size_value)

		image_size_length = 0

		original_image_size_length = original_size_value.slice(original_size_value.index("x") + 2..-1)

		if original_image_size_length.include?('/')

			length_fraction_numerator = original_image_size_length.slice(original_image_size_length.index("/") - 1).to_f
			length_fraction_denominator = original_image_size_length.slice(original_image_size_length.index("/") + 1).to_f
			
			image_size_length += length_fraction_numerator/length_fraction_denominator

		end

		if original_image_size_length.include?('.')
			image_size_length += original_image_size_length[0,5].to_f
		else
			image_size_length += original_image_size_length[0,2].to_f
		end

		return image_size_length

	end

	def compute_poster_size(a, b)

		return a.to_i.to_s + "\"" + "x" + b.to_i.to_s + "\""

	end

	def compute_poster_size_ui(a, b)

		return (a + b).to_i

	end

	def compute_poster_size_category(x)
		
		if (x != 0)

			if (x < 30) 
				return "Petite"
			end

			if (x >= 30 and x <  40)
				return "Small"
			end

			if (x >= 40 and x < 50)
				return "Medium"
			end

			if (x >= 50 and x < 60)
				return "Large"
			end

			if (x >= 60)
				return "Oversize"
			end
		
		end

		return ""

	end


	# GET /generate_template
	# GET /generate_template.json
	def index

		


		$beginning = Time.now

 
		# Load the source Excel file, with all the special products info
		$source = Excel.new("Template_2013_05_10/source.xls")
		#$source = Csv.new("Template_2013_05_10/source.csv")
		$source.default_sheet = $source.sheets.first
		
		# Load the Magento template, which is in Open Office format
		#template = Openoffice.new("http://beta.topart.com/csv/Template_2012_11_28/template.ods")
		#template = Csv.new("Template_2013_05_10/template.csv")
		global_template = Openoffice.new("Template_2013_05_10/template.ods")
		global_template.default_sheet = global_template.sheets.first

		# Automatically scan the template column names and store them in an associative array
		$template_dictionary = Hash.new
		"A".upto("GC") do |alphabet_character|
			cell_content = "#{global_template.cell(1, alphabet_character)}"
			$template_dictionary[cell_content] = alphabet_character
		end

		#global_template = nil

		p "Template headers loaded."

		# Automatically scan the source column names and store them in an associative array
		$source_dictionary = Hash.new
		"A".upto("BU") do |alphabet_character|
			cell_content = "#{$source.cell(1, alphabet_character)}"
			$source_dictionary[cell_content] = alphabet_character
		end

		p "Source headers loaded."

		# Load the retail_material_size spreadsheet file for paper
		$retail_photo_paper = Excel.new("Template_2013_05_10/retail_master.xls")
		$retail_photo_paper.default_sheet = $retail_photo_paper.sheets[0]

		# Load the retail_material_size spreadsheet file for canvas
		$retail_canvas = Excel.new("Template_2013_05_10/retail_master.xls")
		$retail_canvas.default_sheet = $retail_canvas.sheets[2]

		# Load the retail_framing spreadsheet file to extract framing, stretching and matting information
		retail_framing = Excel.new("Template_2013_05_10/retail_master.xls")
		retail_framing.default_sheet = retail_framing.sheets[3]


		# MATERIAL -> PAPER
		# Automatically scan the source column names and store them in an associative array
		$retail_photo_paper_dictionary = Hash.new
		"A".upto("T") do |alphabet_character|
			cell_content = "#{$retail_photo_paper.cell(1, alphabet_character)}"
			$retail_photo_paper_dictionary[cell_content] = alphabet_character
		end

		p "Retail photo paper headers correctly loaded."

		# MATERIAL -> CANVAS
		# Automatically scan the source column names and store them in an associative array
		$retail_canvas_dictionary = Hash.new
		"A".upto("AO") do |alphabet_character|
			cell_content = "#{$retail_canvas.cell(1, alphabet_character)}"
			$retail_canvas_dictionary[cell_content] = alphabet_character
		end

		p "Retail canvas headers correctly loaded."

		$retail_framing_dictionary = Hash.new
		"A".upto("R") do |alphabet_character|
			cell_content = "#{retail_framing.cell(1, alphabet_character)}"
			$retail_framing_dictionary[cell_content] = alphabet_character
		end


		# FRAMING, STRETCHING, MATTING
		# Automatically scan the source column names and store them in an associative array
		# Declare and fill the retail framing table
		$retail_framing_table = Array.new(retail_framing.last_row, 18)
		i = 0

		# Scan all the source rows and process the F21066 items only, and only once at the beginning for efficiency
		2.upto($source.last_row) do |source_line|

			primary_vendor_no = "#{$source.cell(source_line, $source_dictionary["PrimaryVendorNo"])}"

			if primary_vendor_no == "F21066"
				$retail_framing_table[i] = Hash.new

				# Store all the MAS specific fields, which means the majority of them
				"A".upto("R") do |alphabet_character|
					header = "#{retail_framing.cell(1, alphabet_character)}"
					$retail_framing_table[i][header] = "#{$source.cell(source_line, $source_dictionary[header])}"
				end

				# Store the spreadsheet retail prices only
				2.upto(retail_framing.last_row) do |k|
					#$retail_framing_table[i] = Hash.new

					"C".upto("F") do |alphabet_character|
						cell_content = "#{retail_framing.cell(1, alphabet_character)}"

						if $retail_framing_table[i]["Item Code"] == "#{retail_framing.cell(k, $retail_framing_dictionary["Item Code"])}"
							$retail_framing_table[i][cell_content] = "#{retail_framing.cell(k, alphabet_character)}"
						end
					end
				end

				i = i + 1

			end

		end

		p "The F21066 items have been correctly loaded."

		#written_categories = []
		$global_alternate_size_array = Array.new

		# Load a hash table with all the item codes from the products spreadsheet. Used to check the presence of DGs and corresponding posters
		$item_source_line = Hash.new

		#7000.upto($source.last_row) do |source_line|
		2.upto($source.last_row) do |source_line|
			item_code = "#{$source.cell(source_line, $source_dictionary["Item Code"])}"
			$item_source_line[item_code] = source_line
		end

		p "All the source lines have been correctly mapped."

		# We use the following hash table to track DG products that should contain the additional poster size as a custom option
		$posters_and_dgs_hash_table = Hash.new
		$poster_only_hash_table = Hash.new

		$template_counter = 1

		#thread_i = 2
		#row_range = 50
		#limit = $source.last_row
		#limit = 100
		#thread_pool = Array.new

		#while thread_i < limit
		#	thread_pool << Thread.new{parallel_write(thread_i, thread_i + row_range - 1)}
		#	thread_i = thread_i + row_range
		#end	

		
		#for thread in thread_pool
		#	thread.join
		#end

		#$global_alternate_size_array << "XWL4870"

		#temp_i = 2
		#temp_x = 13200

		#while temp_i <= temp_x

		#	item_code = "#{$source.cell(temp_i, $source_dictionary["Item Code"])}"

		#	a = "#{$source.cell(temp_i, $source_dictionary["UDF_ALTS1"])}".gsub(' ','')
		#	b = "#{$source.cell(temp_i, $source_dictionary["UDF_ALTS2"])}".gsub(' ','')
		#	c = "#{$source.cell(temp_i, $source_dictionary["UDF_ALTS3"])}".gsub(' ','')
		#	d = "#{$source.cell(temp_i, $source_dictionary["UDF_ALTS4"])}".gsub(' ','')
			
		#	if !a.blank?
		#		
		#		$global_alternate_size_array << a
		#	end
		#	if !b.blank?
		#	
		#		$global_alternate_size_array << b
		#	end
		#	if !c.blank?
		#		
		#		$global_alternate_size_array << c
		#	end
		#	if !d.blank?
		#		
		#		$global_alternate_size_array << d
		#	end

			# If the current sku is an alternate size of a sku we have already met, then skip it and go to the next item number
		#	if ($global_alternate_size_array.include?(item_code))
		#		
		#		p item_code + " already scanned."

		#		$global_alternate_size_array << item_code
		#	end

		#	temp_i = temp_i + 1

		#end


		t1 = Thread.new{parallel_write(840, $source.last_row)}
		#t1 = Thread.new{parallel_write(2, 10)}
		#t1 = Thread.new{parallel_write(9020, 9030)}
		#t1 = Thread.new{parallel_write(2, 51)}
		#t2 = Thread.new{parallel_write(52, 101)}
		#t3 = Thread.new{parallel_write(102, 151)}
		#t4 = Thread.new{parallel_write(152, 201)}

		t1.join
		#t2.join
		#t3.join
		#t4.join

		puts "The overall running time has been #{Time.now - $beginning} seconds."

		# Accessing this view launch the service automatically
		respond_to do |format|
			format.html # index.html.erb
		end

	end


	def parallel_write(source_line, last_row)

		#RubyProf.start

		#p source_line
		#p last_row

		# Thread specific data
		destination_line = 2
		current_thread_beginning = Time.now
		
		template = Openoffice.new("Template_2013_05_10/template.ods")
		template.default_sheet = template.sheets.first

		while source_line <= last_row

			#p Thread.list.select {|thread| thread.status == "run"}.count
			#loop_start = Time.now
				
			### Fields variables for each product are all assigned here ###
			udf_tar = "#{$source.cell(source_line, $source_dictionary["UDF_TAR"])}"

			# Skip importing items where udf_tar = N
			if udf_tar == "N"

				source_line = source_line + 1
				next
			end

			primary_vendor_no = "#{$source.cell(source_line, $source_dictionary["PrimaryVendorNo"])}"

			# Skip importing the framing related items
			if primary_vendor_no == "F21066"

				source_line = source_line + 1
				next
			end

			item_code = "#{$source.cell(source_line, $source_dictionary["Item Code"])}"
			udf_entity_type = "#{$source.cell(source_line, $source_dictionary["UDF_ENTITYTYPE"])}"

			# If the current sku is an alternate size of a sku we have already met, then skip it and go to the next item number
			if ($global_alternate_size_array.include?(item_code))
				
				p item_code + " already scanned."

				# Assuming the alternate size DG items are the same as the main item number DG item, we skip them as well
				# If the next item in the list is the DG item number, then skip it, otherwise analyze it
				if ($item_source_line[item_code + "DG"] == (source_line + 1))
					source_line = source_line + 2
				else
					source_line = source_line + 1
				end

				# Check if we need to write to csv file now
				if ( item_code == "XWL4870" )

					# Finally, fill the template
					template_file_name = "csv/new_inventory_" + $template_counter.to_s + ".csv"
					p "Creating " + template_file_name + "..."
					template.to_csv(template_file_name)

					puts "The running time for the current .csv file has been #{Time.now - $beginning} seconds."

					$template_counter = $template_counter + 1
					destination_line = 2

					# Reset the template file to store the new rows
					template = Openoffice.new("Template_2013_05_10/template.ods")
					template.default_sheet = template.sheets.first

				end

				next

			end

			# We use this variable to keep track of the right line to take data from.
			scan_line = 0

			if udf_entity_type == "Poster"

				# Compute the correspondig DG item code
				dg_item_code = item_code + "DG"

				# If the poster has a corresponding DG item available
				if $item_source_line[dg_item_code]
					
					$posters_and_dgs_hash_table[dg_item_code] = "true"

					# This will be the line of the corresponding DG item, used for the DG specific attributes only.
					scan_line = $item_source_line[dg_item_code]

				else

					$poster_only_hash_table[item_code] = "true"
					scan_line = source_line

				end
			end

			if udf_entity_type == "Image"

				scan_line = $item_source_line[item_code]

			end


			
			description = "#{$source.cell(source_line, $source_dictionary["Description"])}"
			special_character_index = description.index("^")
			if special_character_index != nil
				#p description
				description = description.gsub("^", "'")
				#p description
			end


			udf_pricecode = "#{$source.cell(source_line, $source_dictionary["UDF_PRICECODE"])}"

			udf_paper_size_cm = "#{$source.cell(source_line, $source_dictionary["UDF_PAPER_SIZE_CM"])}"
			udf_paper_size_in = "#{$source.cell(source_line, $source_dictionary["UDF_PAPER_SIZE_IN"])}"
			udf_image_size_cm = "#{$source.cell(source_line, $source_dictionary["UDF_IMAGE_SIZE_CM"])}"
			udf_image_size_in = "#{$source.cell(source_line, $source_dictionary["UDF_IMAGE_SIZE_IN"])}"


			if udf_paper_size_in.blank? and !udf_paper_size_cm.blank?
				udf_paper_size_in = (compute_image_size_width(udf_paper_size_cm) / 2.54).round(2).to_s + " x " + (compute_image_size_length(udf_paper_size_cm) / 2.54).round(2).to_s
			end

			if udf_image_size_in.blank?

				if !udf_image_size_cm.blank?
					udf_image_size_in = (compute_image_size_width(udf_image_size_cm) / 2.54).round(2).to_s + " x " + (compute_image_size_length(udf_image_size_cm) / 2.54).round(2).to_s
				else
					udf_image_size_in = udf_paper_size_in
				end
			end


			udf_alt_size_1 = "#{$source.cell(source_line, $source_dictionary["UDF_ALTS1"])}".gsub(' ','')
			udf_alt_size_2 = "#{$source.cell(source_line, $source_dictionary["UDF_ALTS2"])}".gsub(' ','')
			udf_alt_size_3 = "#{$source.cell(source_line, $source_dictionary["UDF_ALTS3"])}".gsub(' ','')
			udf_alt_size_4 = "#{$source.cell(source_line, $source_dictionary["UDF_ALTS4"])}".gsub(' ','')


			# Array containing the alternate sizes, to be used later in the code
			alternate_size_array = Array.new
			if !udf_alt_size_1.blank?
				alternate_size_array << udf_alt_size_1
				$global_alternate_size_array << udf_alt_size_1
			end
			if !udf_alt_size_2.blank?
				alternate_size_array << udf_alt_size_2
				$global_alternate_size_array << udf_alt_size_2
			end
			if !udf_alt_size_3.blank?
				alternate_size_array << udf_alt_size_3
				$global_alternate_size_array << udf_alt_size_3
			end
			if !udf_alt_size_4.blank?
				alternate_size_array << udf_alt_size_4
				$global_alternate_size_array << udf_alt_size_4
			end


			udf_oversize = "#{$source.cell(source_line, $source_dictionary["UDF_OVERSIZE"])}"
			udf_serigraph = "#{$source.cell(source_line, $source_dictionary["UDF_SERIGRAPH"])}"
			udf_embossed = "#{$source.cell(source_line, $source_dictionary["UDF_EMBOSSED"])}"
			udf_foil = "#{$source.cell(source_line, $source_dictionary["UDF_FOIL"])}"
			udf_metallic_ink = "#{$source.cell(source_line, $source_dictionary["UDF_METALLICINK"])}"
			udf_specpaper = "#{$source.cell(source_line, $source_dictionary["UDF_SPECPAPER"])}"


			udf_orientation = "#{$source.cell(source_line, $source_dictionary["UDF_ORIENTATION"])}"
			udf_new = "#{$source.cell(source_line, $source_dictionary["UDF_NEW"])}"
			udf_dnd = "#{$source.cell(source_line, $source_dictionary["UDF_DND"])}"
			udf_imsource = "#{$source.cell(source_line, $source_dictionary["UDF_IMSOURCE"])}"

			udf_canvas = "#{$source.cell(scan_line, $source_dictionary["UDF_CANVAS"])}"
			udf_rag = "#{$source.cell(scan_line, $source_dictionary["UDF_RAG"])}"
			udf_photopaper = "#{$source.cell(scan_line, $source_dictionary["UDF_PHOTOPAPER"])}"
			udf_poster = "#{$source.cell(source_line, $source_dictionary["UDF_POSTER"])}"
			
			total_quantity_on_hand = "#{$source.cell(source_line, $source_dictionary["TotalQuantityOnHand"])}".to_i
			udf_decal = "#{$source.cell(scan_line, $source_dictionary["UDF_DECAL"])}"
			udf_embellished = "#{$source.cell(scan_line, $source_dictionary["UDF_EMBELLISHED"])}"
			udf_framed = "#{$source.cell(source_line, $source_dictionary["UDF_FRAMED"])}"


			udf_a4pod = "#{$source.cell(scan_line, $source_dictionary["UDF_A4POD"])}"
			udf_custom_size = "#{$source.cell(scan_line, $source_dictionary["UDF_CUSTOMSIZE"])}"
			udf_petite = "#{$source.cell(scan_line, $source_dictionary["UDF_PETITE"])}"
			udf_small = "#{$source.cell(scan_line, $source_dictionary["UDF_SMALL"])}"
			udf_medium = "#{$source.cell(scan_line, $source_dictionary["UDF_MEDIUM"])}"
			udf_large = "#{$source.cell(scan_line, $source_dictionary["UDF_LARGE"])}"
			udf_osdp = "#{$source.cell(scan_line, $source_dictionary["UDF_OSDP"])}"

			udf_limited = "#{$source.cell(source_line, $source_dictionary["UDF_LIMITED"])}"
			udf_copyright = "#{$source.cell(source_line, $source_dictionary["UDF_COPYRIGHT"])}"
			udf_crline = "#{$source.cell(source_line, $source_dictionary["UDF_CRLINE"])}"
			udf_crimage = "#{$source.cell(source_line, $source_dictionary["UDF_CRIMAGE"])}"
			udf_anycustom = "#{$source.cell(scan_line, $source_dictionary["UDF_ANYCUSTOM"])}"

			udf_maxsfcm = "#{$source.cell(source_line, $source_dictionary["UDF_MAXSFCM"])}"
			udf_maxsfin = "#{$source.cell(source_line, $source_dictionary["UDF_MAXSFIN"])}"

			if !udf_maxsfin.blank?
				udf_maxsfin = udf_maxsfin.to_f
			end


			udf_attributes = "#{$source.cell(source_line, $source_dictionary["UDF_ATTRIBUTES"])}"
			udf_ratio_dec = "#{$source.cell(source_line, $source_dictionary["UDF_RATIODEC"])}".to_f

			udf_largeos = "#{$source.cell(scan_line, $source_dictionary["UDF_LARGEOS"])}"

			suggested_retail_price = "#{$source.cell(source_line, $source_dictionary["SuggestedRetailPrice"])}".to_i
			udf_eco = "#{$source.cell(source_line, $source_dictionary["UDF_ECO"])}"
			udf_fmaxslscm = "#{$source.cell(source_line, $source_dictionary["UDF_FMAXSLSCM"])}"
			udf_fmaxslsin = "#{$source.cell(source_line, $source_dictionary["UDF_FMAXSLSIN"])}"
			udf_fmaxsssin = "#{$source.cell(source_line, $source_dictionary["UDF_FMAXSSSIN"])}"
			udf_fmaxssxcm = "#{$source.cell(source_line, $source_dictionary["UDF_FMAXSSXCM"])}"


			udf_colorcode = "#{$source.cell(source_line, $source_dictionary["UDF_COLORCODE"])}"
			udf_framecat = "#{$source.cell(source_line, $source_dictionary["UDF_FRAMECAT"])}"
			udf_prisubnsubcat = "#{$source.cell(source_line, $source_dictionary["UDF_PRISUBNSUBCAT"])}"
			udf_pricolor = "#{$source.cell(source_line, $source_dictionary["UDF_PRICOLOR"])}"
			udf_pristyle = "#{$source.cell(source_line, $source_dictionary["UDF_PRISTYLE"])}"
			udf_rooms = "#{$source.cell(source_line, $source_dictionary["UDF_ROOMS"])}"


			udf_artshop = "#{$source.cell(source_line, $source_dictionary["UDF_ARTSHOP"])}"
			udf_artshopi = "#{$source.cell(source_line, $source_dictionary["UDF_ARTSHOPI"])}"
			udf_artshopl = "#{$source.cell(source_line, $source_dictionary["UDF_ARTSHOPL"])}"
			udf_nollcavail = "#{$source.cell(source_line, $source_dictionary["UDF_NOLLCAVAIL"])}"
			udf_llcroy = "#{$source.cell(source_line, $source_dictionary["UDF_LLCROY"])}"
			udf_royllcval = "#{$source.cell(source_line, $source_dictionary["UDF_ROYLLCVAL"])}"


			udf_f_m_avail_4_paper = "#{$source.cell(source_line, $source_dictionary["UDF_FMAVAIL4PAPER"])}"
			udf_f_m_avail_4_canvas = "#{$source.cell(source_line, $source_dictionary["UDF_FMAVAIL4CANVAS"])}"
			udf_moulding_width = "#{$source.cell(source_line, $source_dictionary["UDF_MOULDINGWIDTH"])}"
			udf_ratiocode = "#{$source.cell(source_line, $source_dictionary["UDF_RATIOCODE"])}"
			udf_marketattrib = "#{$source.cell(source_line, $source_dictionary["UDF_MARKETATTRIB"])}"
			udf_artist_name = "#{$source.cell(source_line, $source_dictionary["UDF_ARTIST_NAME"])}"

			

			### End of Fields variables assignments ###
			

			template.set(destination_line, $template_dictionary["sku"], item_code)
			template.set(destination_line, $template_dictionary["_attribute_set"], "Topart - Products")
			template.set(destination_line, $template_dictionary["_type"], "simple")

			collections_count = 0

			
			template.set(destination_line + collections_count, $template_dictionary["_category"], "Artists/" + udf_artist_name)
			template.set(destination_line + collections_count, $template_dictionary["_root_category"], "Root Category")
				
			collections_count = collections_count + 1
			


			# Category structure: categories and subcategories
			# Example: x(a;b;c).y.z(f).
			category_array = udf_prisubnsubcat.split(".")

			0.upto(category_array.size-1) do |i|

				open_brace_index = category_array[i].index("(")
				close_brace_index = category_array[i].index(")")
				
				# Category name
				if open_brace_index != nil
					category_name = category_array[i][0..open_brace_index-1].titleize

					# Subcategory list
					subcategory_array = category_array[i][open_brace_index+1..close_brace_index-1].split(";")

					0.upto(subcategory_array.size-1) do |j|

						# This if block is only used once to comput the unique list of categories/subcategories
						#if !written_categories.include?(category_name + "/" + subcategory_array[j].capitalize)
						#if !written_categories.include?(category_name)
							#p category_name + "/" + subcategory_array[j].capitalize
							#written_categories << (category_name + "/" + subcategory_array[j].capitalize)
							#written_categories << (category_name)
						#end

						template.set(destination_line + collections_count, $template_dictionary["_category"], "Subjects/" + category_name + "/" + subcategory_array[j].titleize)
						template.set(destination_line + collections_count, $template_dictionary["_root_category"], "Root Category")

						collections_count = collections_count + 1

					end
				else

					category_name = category_array[i][0..category_array[i].length-1]

					# This if block is only used once to comput the unique list of categories/subcategories
					#if !written_categories.include?(category_name)
						#p category_name
						#written_categories << category_name
					#end

					template.set(destination_line + collections_count, $template_dictionary["_category"], "Subjects/" + category_name)
					template.set(destination_line + collections_count, $template_dictionary["_root_category"], "Root Category")

					collections_count = collections_count + 1

				end

			end


			# Collections
			collections_array = udf_marketattrib.split(".")

			0.upto(collections_array.size-1) do |i|

				collection_name = collections_array[i][0..collections_array[i].length-1]

				template.set(destination_line + collections_count, $template_dictionary["_category"], "Collections/" + collection_name)
				template.set(destination_line + collections_count, $template_dictionary["_root_category"], "Root Category")

				collections_count = collections_count + 1

			end



			template.set(destination_line, $template_dictionary["_product_websites"], "base")
			
			# Alt Size 1, Alt Size 2, Alt Size 3, Alt Size 4
			template.set(destination_line, $template_dictionary["alt_size_1"], udf_alt_size_1)
			template.set(destination_line, $template_dictionary["alt_size_2"], udf_alt_size_2)
			template.set(destination_line, $template_dictionary["alt_size_3"], udf_alt_size_3)
			template.set(destination_line, $template_dictionary["alt_size_4"], udf_alt_size_4)
			
			
			# ItemCodeDesc
			template.set(destination_line, $template_dictionary["description"], description)


			template.set(destination_line, $template_dictionary["enable_googlecheckout"], "1")
			template.set(destination_line, $template_dictionary["udf_orientation"], udf_orientation)


			template.set(destination_line, $template_dictionary["udf_image_size_cm"], udf_image_size_cm)
			template.set(destination_line, $template_dictionary["udf_image_size_in"], udf_image_size_in)
			template.set(destination_line, $template_dictionary["udf_ratiocode"], udf_ratiocode)
			
			template.set(destination_line, $template_dictionary["meta_description"], description)
			keywords_list = udf_attributes.downcase
			template.set(destination_line, $template_dictionary["meta_keyword"], keywords_list)
			template.set(destination_line, $template_dictionary["meta_title"], description)


			template.set(destination_line, $template_dictionary["msrp_display_actual_price_type"], "Use config")
			template.set(destination_line, $template_dictionary["msrp_enabled"], "Use config")

			template.set(destination_line, $template_dictionary["name"], description)
			template.set(destination_line, $template_dictionary["options_container"], "Block after Info Column")

			if udf_oversize == "Y"
				template.set(destination_line, $template_dictionary["udf_oversize"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_oversize"], "No")
			end

			template.set(destination_line, $template_dictionary["udf_paper_size_cm"], udf_paper_size_cm)
			template.set(destination_line, $template_dictionary["udf_paper_size_in"], udf_paper_size_in)

			if udf_a4pod == "Y"
				template.set(destination_line, $template_dictionary["udf_a4pod"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_a4pod"], "No")
			end

			template.set(destination_line, $template_dictionary["price"], "0.0")

			if udf_entity_type == "Image" 
				template.set(destination_line, $template_dictionary["required_options"], "1")
			else
				template.set(destination_line, $template_dictionary["required_options"], "0")
			end

			template.set(destination_line, $template_dictionary["short_description"], description)


			#Status: enabled (1), disabled (2)
			if udf_entity_type == "Image" 
				template.set(destination_line, $template_dictionary["status"], "1")
				template.set(destination_line, $template_dictionary["total_quantity_on_hand"], "0".to_i)
			else
				template.set(destination_line, $template_dictionary["status"], "1")
				template.set(destination_line, $template_dictionary["total_quantity_on_hand"], total_quantity_on_hand)
			end

			template.set(destination_line, $template_dictionary["tax_class_id"], "2")

			#udf_anycustom
			if udf_anycustom == "Y"
				template.set(destination_line, $template_dictionary["udf_anycustom"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_anycustom"], "No")
			end

			#udf_maxsfcm
			template.set(destination_line, $template_dictionary["udf_maxsfcm"], udf_maxsfcm)
			#udf_maxsfin
			template.set(destination_line, $template_dictionary["udf_maxsfin"], udf_maxsfin)

			#udf_largeos
			if udf_largeos == "Y"
				template.set(destination_line, $template_dictionary["udf_largeos"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_largeos"], "No")
			end


			#udf_eco
			if udf_eco == "Y"
				template.set(destination_line, $template_dictionary["udf_eco"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_eco"], "No")
			end

			template.set(destination_line, $template_dictionary["udf_fmaxslscm"], udf_fmaxslscm)
			template.set(destination_line, $template_dictionary["udf_fmaxslsin"], udf_fmaxslsin)
			template.set(destination_line, $template_dictionary["udf_fmaxsssin"], udf_fmaxsssin)
			template.set(destination_line, $template_dictionary["udf_fmaxssxcm"], udf_fmaxssxcm)


			template.set(destination_line, $template_dictionary["udf_colorcode"], udf_colorcode)
			template.set(destination_line, $template_dictionary["udf_framecat"], udf_framecat)
			template.set(destination_line, $template_dictionary["udf_pricolor"], udf_pricolor)
			template.set(destination_line, $template_dictionary["udf_pristyle"], udf_pristyle)
			template.set(destination_line, $template_dictionary["udf_rooms"], udf_rooms)


			template.set(destination_line, $template_dictionary["udf_artshopi"], udf_artshopi)
			template.set(destination_line, $template_dictionary["udf_artshopl"], udf_artshopl)
			template.set(destination_line, $template_dictionary["udf_royllcval"], udf_royllcval)

			if udf_artshop == "Y"
				template.set(destination_line, $template_dictionary["udf_artshop"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_artshop"], "No")
			end

			if udf_nollcavail == "Y"
				template.set(destination_line, $template_dictionary["udf_nollcavail"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_nollcavail"], "No")
			end

			if udf_llcroy == "Y"
				template.set(destination_line, $template_dictionary["udf_llcroy"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_llcroy"], "No")
			end



			if udf_f_m_avail_4_paper == "Y"
				template.set(destination_line, $template_dictionary["udf_f_m_avail_4_paper"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_f_m_avail_4_paper"], "No")
			end

			if udf_f_m_avail_4_canvas == "Y"
				template.set(destination_line, $template_dictionary["udf_f_m_avail_4_canvas"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_f_m_avail_4_canvas"], "No")
			end

			template.set(destination_line, $template_dictionary["udf_moulding_width"], udf_moulding_width)
			template.set(destination_line, $template_dictionary["primary_vendor_no"], primary_vendor_no)


			#udf_canvas
			if udf_canvas == "Y"
				template.set(destination_line, $template_dictionary["udf_canvas"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_canvas"], "No")
			end

			#udf_rag
			if udf_rag == "Y"
				template.set(destination_line, $template_dictionary["udf_rag"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_rag"], "No")
			end

			#udf_photopaper
			if udf_photopaper == "Y"
				template.set(destination_line, $template_dictionary["udf_photo_paper"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_photo_paper"], "No")
			end

			#udf_poster
			if udf_poster == "Y"
				template.set(destination_line, $template_dictionary["udf_poster"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_poster"], "No")
			end

			#udf_decal
			if udf_decal == "Y"
				template.set(destination_line, $template_dictionary["udf_decal"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_decal"], "No")
			end

			#Artist name
			template.set(destination_line, $template_dictionary["udf_artist_name"], udf_artist_name)

			#Copyright
			if udf_copyright == "Y"
				template.set(destination_line, $template_dictionary["udf_copyright"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_copyright"], "No")
			end

			#udf_crimage
			template.set(destination_line, $template_dictionary["udf_crimage"], udf_crimage)

			#udf_crline
			template.set(destination_line, $template_dictionary["udf_crline"], udf_crline)

			#udf_dnd
			if udf_dnd == "Y"
				template.set(destination_line, $template_dictionary["udf_dnd"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_dnd"], "No")
			end

			#udf_embellished
			if udf_embellished == "Y"
				template.set(destination_line, $template_dictionary["udf_embellished"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_embellished"], "No")
			end

			#udf_framed
			if udf_framed == "Y"
				template.set(destination_line, $template_dictionary["udf_framed"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_framed"], "No")
			end

			#udf_imsource
			template.set(destination_line, $template_dictionary["udf_imsource"], udf_imsource)

			#udf_new
			if udf_new == "Y"
				template.set(destination_line, $template_dictionary["udf_new"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_new"], "No")
			end

			#udf_custom_size
			if udf_custom_size == "Y"
				template.set(destination_line, $template_dictionary["udf_custom_size"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_custom_size"], "No")
			end

			#udf_petite
			if udf_petite == "Y"
				template.set(destination_line, $template_dictionary["udf_petite"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_petite"], "No")
			end

			#udf_small
			if udf_small == "Y"
				template.set(destination_line, $template_dictionary["udf_small"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_small"], "No")
			end

			#udf_medium
			if udf_medium == "Y"
				template.set(destination_line, $template_dictionary["udf_medium"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_medium"], "No")
			end

			#udf_large
			if udf_large == "Y"
				template.set(destination_line, $template_dictionary["udf_large"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_large"], "No")
			end

			#udf_osdp
			if udf_osdp == "Y"
				template.set(destination_line, $template_dictionary["udf_osdp"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_osdp"], "No")
			end



			#udf_limited
			if udf_limited == "Y"
				template.set(destination_line, $template_dictionary["udf_limited"], "Yes")
			else
				template.set(destination_line, $template_dictionary["udf_limited"], "No")
			end

			#udf_pricecode
			template.set(destination_line, $template_dictionary["udf_pricecode"], udf_pricecode)

			#udf_ratio_dec
			template.set(destination_line, $template_dictionary["udf_ratiodec"], udf_ratio_dec.to_s)

			template.set(destination_line, $template_dictionary["udf_tar"], "Yes")
			template.set(destination_line, $template_dictionary["status"], "1")

			#URL Key, with the SKU as suffix to keep it unique among products
			template.set(destination_line, $template_dictionary["url_key"], description.gsub(/[ ]/, '-')  << "-" << item_code)

			template.set(destination_line, $template_dictionary["visibility"], "4")
			template.set(destination_line, $template_dictionary["weight"], "1")

			

	
			template.set(destination_line, $template_dictionary["min_qty"], "0")
			template.set(destination_line, $template_dictionary["use_config_min_qty"], "1")
			template.set(destination_line, $template_dictionary["is_qty_decimal"], "0")
			template.set(destination_line, $template_dictionary["backorders"], "0") 

			template.set(destination_line, $template_dictionary["use_config_backorders"], "1")
			template.set(destination_line, $template_dictionary["min_sale_qty"], "1")
			template.set(destination_line, $template_dictionary["use_config_min_sale_qty"], "1")
			template.set(destination_line, $template_dictionary["max_sale_qty"], "0")
			template.set(destination_line, $template_dictionary["use_config_max_sale_qty"], "1")




			if udf_entity_type == "Image"
				template.set(destination_line, $template_dictionary["is_in_stock"], "0")
				template.set(destination_line, $template_dictionary["use_config_notify_stock_qty"], "0")
				template.set(destination_line, $template_dictionary["manage_stock"], "0")
				template.set(destination_line, $template_dictionary["use_config_manage_stock"], "0")
				template.set(destination_line, $template_dictionary["qty"], "0")
				template.set(destination_line, $template_dictionary["has_options"], "1")
			else
				template.set(destination_line, $template_dictionary["is_in_stock"], "1")
				template.set(destination_line, $template_dictionary["use_config_notify_stock_qty"], "1")
				template.set(destination_line, $template_dictionary["manage_stock"], "0")
				template.set(destination_line, $template_dictionary["use_config_manage_stock"], "0")
				template.set(destination_line, $template_dictionary["qty"], total_quantity_on_hand)
				template.set(destination_line, $template_dictionary["has_options"], "0")
			end



			
			template.set(destination_line, $template_dictionary["stock_status_changed_auto"], "0")				
			template.set(destination_line, $template_dictionary["use_config_qty_increments"], "1")
			template.set(destination_line, $template_dictionary["qty_increments"], "0")
			template.set(destination_line, $template_dictionary["use_config_enable_qty_inc"], "1")
			template.set(destination_line, $template_dictionary["enable_qty_increments"], "0") 
			template.set(destination_line, $template_dictionary["is_decimal_divided"], "0")


			


			if udf_entity_type == "Poster" and ( ((udf_imsource == "San Diego" || udf_imsource == "Italy") and total_quantity_on_hand > -1) || udf_imsource == "Old World")

				image_size_width = compute_image_size_width(udf_image_size_in)
				image_size_length = compute_image_size_length(udf_image_size_in)

				poster_size_ui = compute_poster_size_ui(image_size_width, image_size_length)
				poster_size = compute_poster_size(image_size_width, image_size_length)

				template.set(destination_line, $template_dictionary["size_category"], compute_poster_size_category(poster_size_ui))


				# Embellishments
				if udf_metallic_ink == "Y"
					template.set(destination_line, $template_dictionary["udf_metallic_ink"], "Yes")
				else
					template.set(destination_line, $template_dictionary["udf_metallic_ink"], "No")
				end
				if udf_foil == "Y"
					template.set(destination_line, $template_dictionary["udf_foil"], "Yes")
				else
					template.set(destination_line, $template_dictionary["udf_foil"], "No")
				end
				if udf_serigraph == "Y"
					template.set(destination_line, $template_dictionary["udf_serigraph"], "Yes")
				else
					template.set(destination_line, $template_dictionary["udf_serigraph"], "No")
				end
				if udf_embossed == "Y"
					template.set(destination_line, $template_dictionary["udf_embossed"], "Yes")
				else
					template.set(destination_line, $template_dictionary["udf_embossed"], "No")
				end
				if udf_specpaper == "Y"
					template.set(destination_line, $template_dictionary["udf_specpaper"], "Yes")
				else
					template.set(destination_line, $template_dictionary["udf_specpaper"], "No")
				end

			end


			########## Custom options columns ##########

			# MATERIAL: paper and canvas are static hard-coded options.

			########### Material ###############

			template.set(destination_line, $template_dictionary["_custom_option_type"], "radio")
			template.set(destination_line, $template_dictionary["_custom_option_title"], "Material")
			template.set(destination_line, $template_dictionary["_custom_option_is_required"], "1")
			template.set(destination_line, $template_dictionary["_custom_option_max_characters"], "0")
			template.set(destination_line, $template_dictionary["_custom_option_sort_order"], "0")


			# Each material option is displayed according to the corresponding udf values
			if udf_entity_type == "Poster" and ( ((udf_imsource == "San Diego" || udf_imsource == "Italy") and total_quantity_on_hand > -1) || udf_imsource == "Old World")

				template.set(destination_line, $template_dictionary["_custom_option_row_title"], "Poster Paper")
				template.set(destination_line, $template_dictionary["_custom_option_row_price"], "0.00")
				template.set(destination_line, $template_dictionary["_custom_option_row_sku"], "material_posterpaper")
				template.set(destination_line, $template_dictionary["_custom_option_row_sort"], "0")

				destination_line = destination_line + 1

			end


			
			if udf_photopaper == "Y"

				template.set(destination_line, $template_dictionary["_custom_option_row_title"], "Photo Paper")
				template.set(destination_line, $template_dictionary["_custom_option_row_price"], "0.00")
				template.set(destination_line, $template_dictionary["_custom_option_row_sku"], "material_photopaper")
				template.set(destination_line, $template_dictionary["_custom_option_row_sort"], "1")

				destination_line = destination_line + 1

			end

			# If not available as poster only
			if udf_canvas == "Y"

				template.set(destination_line, $template_dictionary["_custom_option_row_title"], "Canvas")
				template.set(destination_line, $template_dictionary["_custom_option_row_price"], "0.00")
				template.set(destination_line, $template_dictionary["_custom_option_row_sku"], "material_canvas")
				template.set(destination_line, $template_dictionary["_custom_option_row_sort"], "2")

				destination_line = destination_line + 1

			end

			########### End of Material ###############


			#############SIZE#############
			template.set(destination_line, $template_dictionary["_custom_option_type"], "radio")
			template.set(destination_line, $template_dictionary["_custom_option_title"], "Size")
			template.set(destination_line, $template_dictionary["_custom_option_is_required"], "1")
			template.set(destination_line, $template_dictionary["_custom_option_max_characters"], "0")
			template.set(destination_line, $template_dictionary["_custom_option_sort_order"], "1")
			
			# We need to extract the right prices, looking them up by (i.e. matching) the ratio column

			# Extract and map the border treatments:
			# 1) Scan for every row into the master paper and master canvas sheets
			# 2) check if the ratio matches the one contained in the product attribute 
			# 3) If the 2 ratios match, then copy the specific retail price option

			match_index = 0


			########## IF POSTER IS IN STOCK ####################
			# Change the minimum total quantity on hand when it is ready in MAS, from -1 to 0
			# The poster is available only when it is in stock
			if udf_entity_type == "Poster" and ( ((udf_imsource == "San Diego" || udf_imsource == "Italy") and total_quantity_on_hand > -1) || udf_imsource == "Old World")


				size_name = "Poster Paper"

				image_size_width = compute_image_size_width(udf_image_size_in)
				image_size_length = compute_image_size_length(udf_image_size_in)

				poster_size = compute_poster_size(image_size_width, image_size_length)
				poster_size_ui = compute_poster_size_ui(image_size_width, image_size_length)


				template.set(destination_line, $template_dictionary["_custom_option_row_title"], size_name + ": " + poster_size)
				if suggested_retail_price != 0
					template.set(destination_line, $template_dictionary["_custom_option_row_price"], suggested_retail_price)
				else
					template.set(destination_line, $template_dictionary["_custom_option_row_price"], "0.0")
				end

				size_category = compute_poster_size_category(poster_size_ui).downcase

				template.set(destination_line, $template_dictionary["_custom_option_row_sku"], "size_posterpaper_" + size_category + "_ui_" + poster_size_ui.to_i.to_s + "_width_" + image_size_width.to_i.to_s + "_length_" + image_size_length.to_i.to_s)
				template.set(destination_line, $template_dictionary["_custom_option_row_sort"], match_index)

				destination_line = destination_line + 1

				match_index = match_index + 1



				# Extract the alternate sizes here
				if !alternate_size_array.empty?

					for i_th_alt_size in alternate_size_array

						alternate_size_line = $item_source_line[i_th_alt_size]
						alternate_size = "#{$source.cell(alternate_size_line, $source_dictionary["UDF_IMAGE_SIZE_IN"])}"

						# Alternate size parameters: to be passed later in a dedicated function
						size_name = "Poster Paper"
						
						image_size_width = compute_image_size_width(alternate_size)
						image_size_length = compute_image_size_length(alternate_size)

						poster_size = compute_poster_size(image_size_width, image_size_length)
						poster_size_ui = compute_poster_size_ui(image_size_width, image_size_length)
						
						suggested_retail_price = "#{$source.cell(alternate_size_line, $source_dictionary["SuggestedRetailPrice"])}".to_i

						size_category = compute_poster_size_category(poster_size_ui).downcase

						template.set(destination_line, $template_dictionary["_custom_option_row_title"], size_name + ": " + poster_size)
						if suggested_retail_price != 0
							template.set(destination_line, $template_dictionary["_custom_option_row_price"], suggested_retail_price)
						else
							template.set(destination_line, $template_dictionary["_custom_option_row_price"], "0.0")
						end
						template.set(destination_line, $template_dictionary["_custom_option_row_sku"], "size_posterpaper_" + size_category + "_altsize_" + i_th_alt_size.downcase + "_ui_" + poster_size_ui.to_i.to_s + "_width_" + image_size_width.to_i.to_s + "_length_" + image_size_length.to_i.to_s)
						template.set(destination_line, $template_dictionary["_custom_option_row_sort"], match_index)

						destination_line = destination_line + 1

						match_index = match_index + 1

					end

				end

			end

			########## end of IF POSTER IS IN STOCK ####################


			########## IF UDF_PHOTOPAPER == Y ####################
			if udf_photopaper == "Y"

				# If not available as poster only
				if $poster_only_hash_table[item_code] != "true"

					custom_size_ui_to_skip = 0
					min_delta = 1000;

					# First pass: scan all the available UI sizes
					2.upto($retail_photo_paper.last_row) do |i|

						retail_ratio_dec = "#{$retail_photo_paper.cell(i, $retail_photo_paper_dictionary["Decimal Ratio"])}".to_f

						if udf_ratio_dec == retail_ratio_dec

							size_paper_ui = "#{$retail_photo_paper.cell(i, $retail_photo_paper_dictionary["UI"])}".to_i

							delta = poster_size_ui - size_paper_ui
							delta = delta.abs

							if delta < min_delta
								custom_size_ui_to_skip = size_paper_ui
								min_delta = delta
							end
						end

					end

					# Master Photo Paper Sheet
					2.upto($retail_photo_paper.last_row) do |i|

						retail_ratio_dec = "#{$retail_photo_paper.cell(i, $retail_photo_paper_dictionary["Decimal Ratio"])}".to_f
						size_photopaper_ui = "#{$retail_photo_paper.cell(i, $retail_photo_paper_dictionary["UI"])}".to_i
						image_source = "#{$retail_photo_paper.cell(i, $retail_photo_paper_dictionary["Image Source"])}"

						image_length = "#{$retail_photo_paper.cell(i, $retail_photo_paper_dictionary["Length"])}".to_f
						image_width = "#{$retail_photo_paper.cell(i, $retail_photo_paper_dictionary["Width"])}".to_f

						short_side = 0
						if image_length < image_width
							short_side = image_length
						else
							short_side = image_width
						end


						# Check for available sizes: the poster size replaces the closes photo paper digital size
						if udf_ratio_dec == retail_ratio_dec and size_photopaper_ui != custom_size_ui_to_skip and udf_imsource == image_source and (udf_maxsfin.blank? or short_side <= udf_maxsfin)

							#retail_ratio_dec = "#{$retail_photo_paper.cell(i, $retail_photo_paper_dictionary["Decimal Ratio"])}"
							size_name = "#{$retail_photo_paper.cell(i, $retail_photo_paper_dictionary["Size Description"])}"

							allowed_size = "false"
							
							# Match the right sizes
							if (udf_petite == "Y" && size_name == "Petite") || (udf_small == "Y" && size_name == "Small") || (udf_medium == "Y" && size_name == "Medium") || (udf_large == "Y" && size_name == "Large") || (udf_osdp == "Y" && size_name == "Oversize") || (udf_largeos == "Y" && size_name == "Oversize Large")

								allowed_size = "true"

							end

							# If the size is allowed, then create the corresponding option
							if allowed_size == "true"

								size_price = "#{$retail_photo_paper.cell(i, $retail_photo_paper_dictionary["Rolled Paper - Estimated Retail"])}" 
								size_photopaper_length = "#{$retail_photo_paper.cell(i, $retail_photo_paper_dictionary["Length"])}".to_i.to_s
								size_photopaper_width = "#{$retail_photo_paper.cell(i, $retail_photo_paper_dictionary["Width"])}".to_i.to_s

								template.set(destination_line, $template_dictionary["_custom_option_row_title"], size_name + ": " + size_photopaper_width + "\""  + "x" + size_photopaper_length + "\"")
								template.set(destination_line, $template_dictionary["_custom_option_row_price"], size_price)

								if size_name.downcase == "oversize large"
									size_name = "Oversize_Large"
								end

				
								template.set(destination_line, $template_dictionary["_custom_option_row_sku"], "size_photopaper_" + size_name.downcase + "_ui_" + size_photopaper_ui.to_s + "_width_" + size_photopaper_width.to_s + "_length_" + size_photopaper_length.to_s)
								template.set(destination_line, $template_dictionary["_custom_option_row_sort"], match_index)

								destination_line = destination_line + 1

								match_index = match_index + 1

							end

						end

					end

				end
			end

			########## end IF UDF_PHOTOPAPER == Y ####################


			########## IF UDF_CANVAS == Y ####################
			if udf_canvas == "Y"

				# Master Canvas Sheet
				2.upto($retail_canvas.last_row) do |i|

					retail_ratio_dec = "#{$retail_canvas.cell(i, $retail_canvas_dictionary["Decimal Ratio"])}".to_f
					image_source = "#{$retail_photo_paper.cell(i, $retail_photo_paper_dictionary["Image Source"])}"
					image_length = "#{$retail_photo_paper.cell(i, $retail_photo_paper_dictionary["Length"])}".to_f
					image_width = "#{$retail_photo_paper.cell(i, $retail_photo_paper_dictionary["Width"])}".to_f

					short_side = 0
					if image_length < image_width
						short_side = image_length
					else
						short_side = image_width
					end
					
					count = 0

					# Check for available sizes and border treatments prices
					if udf_ratio_dec == retail_ratio_dec and udf_imsource == image_source and (udf_maxsfin.blank? or short_side <= udf_maxsfin)

						size_name = "#{$retail_canvas.cell(i, $retail_canvas_dictionary["Size Description"])}"	

						allowed_size = "false"
							
						# Match the right sizes
						if (udf_petite == "Y" && size_name == "Petite") || (udf_small == "Y" && size_name == "Small") || (udf_medium == "Y" && size_name == "Medium") || (udf_large == "Y" && size_name == "Large") || (udf_osdp == "Y" && size_name == "Oversize") || (udf_largeos == "Y" && size_name == "Oversize Large")

							allowed_size = "true"

						end

						# If the size is allowed, then create the corresponding option
						if allowed_size == "true"

							size_price_treatment_1 = "#{$retail_canvas.cell(i, $retail_canvas_dictionary["Rolled Canvas White Border -  Estimated Retail"])}"
							size_price_treatment_2 = "#{$retail_canvas.cell(i, $retail_canvas_dictionary['Rolled Canvas 2" Black Border - Estimated Retail'])}"
							size_price_treatment_3 = "#{$retail_canvas.cell(i, $retail_canvas_dictionary['Rolled Canvas 2" Mirror Border -  Estimated Retail'])}"

							size_canvas_length = "#{$retail_canvas.cell(i, $retail_canvas_dictionary["Length"])}".to_i.to_s
							size_canvas_width = "#{$retail_canvas.cell(i, $retail_canvas_dictionary["Width"])}".to_i.to_s
							
							size_prices = Array.new
							size_prices << size_price_treatment_1 << size_price_treatment_2 << size_price_treatment_3
							size_canvas_ui = "#{$retail_canvas.cell(i, $retail_canvas_dictionary["UI"])}".to_i


							0.upto(2) do |count|

								template.set(destination_line, $template_dictionary["_custom_option_row_title"], size_name + ": " + size_canvas_width + "\""  + "x" + size_canvas_length + "\"")
								template.set(destination_line, $template_dictionary["_custom_option_row_price"], size_prices[count])

								if size_name.downcase == "oversize large"
									size_name = "Oversize_Large"
								end

								#_custom_option_row_sku
								template.set(destination_line, $template_dictionary["_custom_option_row_sku"], "size_canvas_" + size_name.downcase + "_treatment_" + (count+1).to_s + "_ui_" + size_canvas_ui.to_s + "_width_" + size_canvas_width.to_s + "_length_" + size_canvas_length.to_s)
								#_custom_option_row_sort
								template.set(destination_line, $template_dictionary["_custom_option_row_sort"], match_index + count)

								destination_line = destination_line + 1

								count = count
							
							end

							match_index = match_index + 1 + count

						end

					end

				end


				# BORDER TREATMENTS for canvas
				# If not available as poster only
				if $poster_only_hash_table[item_code] != "true"

					########### Border Treatments ###############
					# Border Treatments and Stretching options (including names) are static

					template.set(destination_line, $template_dictionary["_custom_option_type"], "radio")
					template.set(destination_line, $template_dictionary["_custom_option_title"], "Borders")
					template.set(destination_line, $template_dictionary["_custom_option_is_required"], "1")
					template.set(destination_line, $template_dictionary["_custom_option_max_characters"], "0")
					template.set(destination_line, $template_dictionary["_custom_option_sort_order"], "2")
					
					template.set(destination_line, $template_dictionary["_custom_option_row_title"], "None")
					template.set(destination_line, $template_dictionary["_custom_option_row_price"], "0.0")
					template.set(destination_line, $template_dictionary["_custom_option_row_sku"], "treatments_none")
					template.set(destination_line, $template_dictionary["_custom_option_row_sort"], "0")

					destination_line = destination_line + 1
					

					template.set(destination_line, $template_dictionary["_custom_option_row_title"], "3\" White Border")
					template.set(destination_line, $template_dictionary["_custom_option_row_price"], "0.0")
					template.set(destination_line, $template_dictionary["_custom_option_row_sku"], "border_treatment_3_inches_of_white")
					template.set(destination_line, $template_dictionary["_custom_option_row_sort"], "1")

					destination_line = destination_line + 1


					template.set(destination_line, $template_dictionary["_custom_option_row_title"], "2\" Black Border + 1\" White")
					template.set(destination_line, $template_dictionary["_custom_option_row_price"], "0.0") 
					template.set(destination_line, $template_dictionary["_custom_option_row_sku"], "border_treatment_2_inches_of_black_and_1_inch_of_white")
					template.set(destination_line, $template_dictionary["_custom_option_row_sort"], "2")

					destination_line = destination_line + 1

					template.set(destination_line, $template_dictionary["_custom_option_row_title"], "2\" Mirrored Border + 1\" White")
					template.set(destination_line, $template_dictionary["_custom_option_row_price"], "0.0")
					template.set(destination_line, $template_dictionary["_custom_option_row_sku"], "border_treatment_2_inches_mirrored_and_1_inch_of_white")
					template.set(destination_line, $template_dictionary["_custom_option_row_sort"], "3")

					destination_line = destination_line + 1

				end


				########### Canvas Stretching ###############
				0.upto($retail_framing_table.length - 2) do |i|

					udf_entity_type = $retail_framing_table[i]["UDF_ENTITYTYPE"]

					if udf_entity_type == "Stretch"

						stretch_item_number = $retail_framing_table[i]["Item Code"].downcase
						stretch_ui_price = $retail_framing_table[i]["United Inch TAR Retail"]

						template.set(destination_line, $template_dictionary["_custom_option_type"], "checkbox")
						template.set(destination_line, $template_dictionary["_custom_option_title"], "Canvas Stretching")
						template.set(destination_line, $template_dictionary["_custom_option_is_required"], "0")
						template.set(destination_line, $template_dictionary["_custom_option_max_characters"], "0")
						template.set(destination_line, $template_dictionary["_custom_option_sort_order"], "3")
						
						stretching_index = 0

						template.set(destination_line, $template_dictionary["_custom_option_row_title"], "1.5\" Gallery Wrap Stretching")
						template.set(destination_line, $template_dictionary["_custom_option_row_price"], stretch_ui_price.to_s) 
						template.set(destination_line, $template_dictionary["_custom_option_row_sku"], stretch_item_number)
						template.set(destination_line, $template_dictionary["_custom_option_row_sort"], stretching_index)

						destination_line = destination_line + 1
						stretching_index = stretching_index + 1

					end
				end
			
			end
			########## end of IF UDF_CANVAS == Y ####################





			########### FRAMING ###########
			
			########## if UDF_FRAMED == Y ####################
			if udf_framed == "Y"

				template.set(destination_line, $template_dictionary["_custom_option_type"], "drop_down")
				template.set(destination_line, $template_dictionary["_custom_option_title"], "Frame")
				template.set(destination_line, $template_dictionary["_custom_option_is_required"], "1") 
				template.set(destination_line, $template_dictionary["_custom_option_max_characters"], "0")
				template.set(destination_line, $template_dictionary["_custom_option_sort_order"], "4")

				frame_count = 0;
				mats_count = 0;

				# Add the No Frame option
				template.set(destination_line, $template_dictionary["_custom_option_row_sku"], "frame_none")
				template.set(destination_line, $template_dictionary["_custom_option_row_title"], "No Frame")
				template.set(destination_line, $template_dictionary["_custom_option_row_price"], "0.0")
				template.set(destination_line, $template_dictionary["_custom_option_row_sort"], frame_count)

				destination_line = destination_line + 1
				frame_count = frame_count + 1

				# Only scan the framing options
				0.upto($retail_framing_table.length - 2) do |i|
						
					udf_entity_type = $retail_framing_table[i]["UDF_ENTITYTYPE"]
					
					if udf_entity_type == "Frame"

						frame_name = $retail_framing_table[i]["Description"]

						frame_item_number = $retail_framing_table[i]["Item Code"].downcase
						frame_ui_price = $retail_framing_table[i]["United Inch TAR Retail"]
						frame_flat_mounting_price = $retail_framing_table[i]["Flat Mounting Cost"]

						frame_for_paper = $retail_framing_table[i]["UDF_FMAVAIL4PAPER"]
						frame_for_canvas = $retail_framing_table[i]["UDF_FMAVAIL4CANVAS"]

						# Scan the category names and add each of them to an array, used to add it only once
						category_name = $retail_framing_table[i]["UDF_FRAMECAT"].downcase
				

						# Each framing option has a different price for each size (UI) available

						# Available for Paper
						if frame_for_paper == "Y"

							template.set(destination_line, $template_dictionary["_custom_option_row_sku"], frame_item_number)
							template.set(destination_line, $template_dictionary["_custom_option_row_title"], frame_name)
							template.set(destination_line, $template_dictionary["_custom_option_row_price"], frame_ui_price.to_s)
							template.set(destination_line, $template_dictionary["_custom_option_row_sort"], frame_count)

							destination_line = destination_line + 1
							frame_count = frame_count + 1

						end

						# Available for Canvas
						if frame_for_canvas == "Y"

							template.set(destination_line, $template_dictionary["_custom_option_row_sku"], frame_item_number)
							template.set(destination_line, $template_dictionary["_custom_option_row_title"], frame_name)
							template.set(destination_line, $template_dictionary["_custom_option_row_price"], frame_ui_price.to_s)
							template.set(destination_line, $template_dictionary["_custom_option_row_sort"], frame_count)

							destination_line = destination_line + 1
							frame_count = frame_count + 1

						end

					end


				end


				


				########### MATTING ###########
				template.set(destination_line, $template_dictionary["_custom_option_type"], "radio")
				template.set(destination_line, $template_dictionary["_custom_option_title"], "Mats")
				template.set(destination_line, $template_dictionary["_custom_option_is_required"], "1") 
				template.set(destination_line, $template_dictionary["_custom_option_max_characters"], "0")
				template.set(destination_line, $template_dictionary["_custom_option_sort_order"], "5")



				0.upto($retail_framing_table.length - 2) do |i|

					udf_entity_type = $retail_framing_table[i]["UDF_ENTITYTYPE"]

					if udf_entity_type == "Mat"

						mat_name = $retail_framing_table[i]["Description"]
						mat_item_number = $retail_framing_table[i]["Item Code"].downcase 

						mat_ui_price = $retail_framing_table[i]["United Inch TAR Retail"]
						mats_for_paper = $retail_framing_table[i]["UDF_FMAVAIL4PAPER"]
						mats_for_canvas = $retail_framing_table[i]["UDF_FMAVAIL4CANVAS"]
						mats_color = $retail_framing_table[i]["UDF_COLORCODE"]
						category_name = $retail_framing_table[i]["UDF_FRAMECAT"].downcase


						# Available for Paper
						if mats_for_paper == "Y"

							template.set(destination_line, $template_dictionary["_custom_option_row_sku"], mat_item_number)
							template.set(destination_line, $template_dictionary["_custom_option_row_title"], mat_name)
							template.set(destination_line, $template_dictionary["_custom_option_row_price"], mat_ui_price.to_s)
							template.set(destination_line, $template_dictionary["_custom_option_row_sort"], mats_count)

							destination_line = destination_line + 1
							mats_count = mats_count + 1
						end

					end
				end

				template.set(destination_line, $template_dictionary["_custom_option_row_sku"], "mats_none")
				template.set(destination_line, $template_dictionary["_custom_option_row_title"], "No Mats")
				template.set(destination_line, $template_dictionary["_custom_option_row_price"], "0.0")
				template.set(destination_line, $template_dictionary["_custom_option_row_sort"], mats_count)

				destination_line = destination_line + 1
				mats_count = mats_count + 1

			end
			########## end of if UDF_FRAMED == Y ####################			


			####### CUSTOM SIZE: HEIGHT #########
			template_column = $template_dictionary["_custom_option_type"]
			template.set(destination_line, template_column, "field")
			#_custom_option_title
			template_column = $template_dictionary["_custom_option_title"]
			template.set(destination_line, template_column, "Height")
			#_custom_option_is_required
			template_column = $template_dictionary["_custom_option_is_required"]
			template.set(destination_line, template_column, "0")
			#_custom_option_max_characters
			template_column = $template_dictionary["_custom_option_max_characters"]
			template.set(destination_line, template_column, "0")
			#_custom_option_sort_order
			template_column = $template_dictionary["_custom_option_sort_order"]
			template.set(destination_line, template_column, "6")

			destination_line = destination_line + 1

			####### CUSTOM SIZE: WIDTH #########
			template_column = $template_dictionary["_custom_option_type"]
			template.set(destination_line, template_column, "field")
			#_custom_option_title
			template_column = $template_dictionary["_custom_option_title"]
			template.set(destination_line, template_column, "Width")
			#_custom_option_is_required
			template_column = $template_dictionary["_custom_option_is_required"]
			template.set(destination_line, template_column, "0")
			#_custom_option_max_characters
			template_column = $template_dictionary["_custom_option_max_characters"]
			template.set(destination_line, template_column, "0")
			#_custom_option_sort_order
			template_column = $template_dictionary["_custom_option_sort_order"]
			template.set(destination_line, template_column, "7")

			destination_line = destination_line + 1	
			
			
			# Compute the maximum count among all the multi select options
			# then add it to the destination line count for the next product to be written
			custom_options_array_size = 0

			multi_select_options = Array.new
			multi_select_options << collections_count

			if udf_entity_type == "Image"
				multi_select_options << custom_options_array_size
			end

			max_count =  multi_select_options.max
			
			# Increase the destination line to the correct number
			destination_line = destination_line + max_count
			destination_line = destination_line + 1


			p source_line.to_s + "/" + $source.last_row.to_s

			if ( ( source_line % 200 == 0 or ((source_line + 1) % 200 == 0) ) or source_line == last_row - 1 )

				# Finally, fill the template
				template_file_name = "csv/new_inventory_" + $template_counter.to_s + ".csv"
				p "Creating " + template_file_name + "..."
				template.to_csv(template_file_name)

				puts "The running time for the current .csv file has been #{Time.now - $beginning} seconds."

				$template_counter = $template_counter + 1
				destination_line = 2

				# Reset the template file to store the new rows
				template = Openoffice.new("Template_2013_05_10/template.ods")
				template.default_sheet = template.sheets.first
			end

			source_line = scan_line + 1


			#loop_end = Time.now
			#loop_time = loop_end - loop_start
			#p "The loop running time is #{loop_time} seconds."
			#puts "The running time for thread " + Thread.current.object_id.to_s + " has been #{Time.now - current_thread_beginning} seconds."

		end

		# Finally, fill the template
		#template_file_name = "csv/new_inventory_" + Thread.current.object_id.to_s + ".csv"
		#p "Creating " + template_file_name + "..."
		#template.to_csv(template_file_name)


		#result = RubyProf.stop
		#File.open "#{Rails.root}/tmp/running_graph.html", 'w' do |file|
		#	RubyProf::GraphHtmlPrinter.new(result).print(file)
		#end

	end

end
