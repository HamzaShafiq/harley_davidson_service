require 'bundler'
Bundler.setup
require 'savon'
require "yaml"
require 'roo'

class Load

  def initialize()
    # Initialize SOAP client using the WSDL
    config = YAML.load(File.open('config/config.yml').read)
    @client = Savon.client(:wsdl => config["wsdl"], log: false, ssl_verify_mode: :none)
  end

  # Call Load on the EchoService
  def load_request()
    config = YAML.load(File.open('config/config.yml').read)

    xlsx = Roo::Spreadsheet.open(config["path_to_file"])

    xlsx.each_with_pagename do |name, sheet|
      sheet.each_row_streaming(pad_cells: true, offset: 1) do |row|
        next if row.eql?(sheet.first_row)
        response = @client.call(:load_mnf_mdl, message: { xmlDocument: generate_xml(row) })
        File.open('log/responses.log','a') do |line|
           line.puts "\r" + "Response for #{row[0]} from LeasePak: #{response}"
        end
      end
    end
  rescue Savon::Error => exception
    File.open('log/errors.log','a') do |line|
       line.puts "\r" + "An error occurred while calling the LeasePak: #{exception.message}"
    end
  end

  private

  def generate_xml(row)
    xml = Builder::XmlMarkup.new

    xml.LPAuxTableMaint do |table|
      table.RECORD do |d|
        d.manf(row[0])
        d.model(row[1])
        d.name(row[2])
      end
    end
  end
end

if __FILE__ == $0
  # Initialize the EchoService client and call operations
  load = Load.new
  load.load_request()
end
