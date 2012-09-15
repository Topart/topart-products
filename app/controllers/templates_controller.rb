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
		
		2.upto(5) do |line|
		
			template.set(line, 1, "#{source.cell(line,'A')}")
	  
		end
		
		template.to_csv("filled_template.csv")
 
		respond_to do |format|
			format.html # index.html.erb
		end

	end

end
