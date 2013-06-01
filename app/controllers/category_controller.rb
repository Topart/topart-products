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
			
		
		@written_categories = []


		source_line = 2
		while source_line <= source.last_row
		#while source_line <= 31
				
			
			@udf_prisubnsubcat = "#{source.cell(source_line, "BF")}"
			

			# Category structure: categories and subcategories
			# Example: x(a;b;c).y.z(f).
			@category_array = @udf_prisubnsubcat.split(".")

			0.upto(@category_array.size-1) do |i|

				@open_brace_index = @category_array[i].index("(")
				@close_brace_index = @category_array[i].index(")")
				
				# Category name
				if @open_brace_index != nil
					@category_name = @category_array[i][0..@open_brace_index-1]

					# Subcategory list
					@subcategory_array = @category_array[i][@open_brace_index+1..@close_brace_index-1].split(";")

					0.upto(@subcategory_array.size-1) do |j|

						# This if block is only used once to comput the unique list of categories/subcategories
						#if !@written_categories.include?(@category_name + "/" + @subcategory_array[j].capitalize)
						if !@written_categories.include?(@category_name)
						#if !@written_categories.include?(@subcategory_array[j])
							#p @category_name + "/" + @subcategory_array[j].capitalize
							#@written_categories << (@category_name + "/" + @subcategory_array[j].capitalize)
							@written_categories << (@category_name)
							#@written_categories << (@subcategory_array[j])

							#p @category_name + " -" + source_line.to_s
						end

					end
				else

					@category_name = @category_array[i][0..@category_array[i].length-1]

					# This if block is only used once to comput the unique list of categories/subcategories
					if !@written_categories.include?(@category_name)
						#p @category_name
						@written_categories << @category_name

						#p @category_name + " - " + source_line.to_s
					end

				end

			end

			source_line = source_line + 1

		end


		# Write the categories to a CSV file for further processing
		category_list = Openoffice.new("Template_2013_05_10/category_list.ods")
		category_list.default_sheet = category_list.sheets.first

		@written_categories.sort!
		@category_counter = 1	
  		
  		@written_categories.each do |row|
  			category_list.set(@category_counter, "A", "7")
  			category_list.set(@category_counter, "B", row)

			@category_counter = @category_counter + 1 
		end

		category_list.to_csv("top_level_categories.csv")


		#@written_categories.sort!
		@unique_counter = 1		
		File.open("categories.txt", "w") do |f|
  			@written_categories.each do |row| f << @unique_counter << ") " << row << "\n" 
				@unique_counter = @unique_counter + 1 
			end
		end

		# Accessing this view launch the service automatically
		respond_to do |format|
			format.html # index.html.erb
		end

	end

end
