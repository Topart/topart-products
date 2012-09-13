require 'rubygems'
require 'roo'

class TemplatesController < ApplicationController

	# GET /generate_template
	# GET /generate_template.json
	def index
 
		respond_to do |format|
			format.html # index.html.erb
		end

	end

end
