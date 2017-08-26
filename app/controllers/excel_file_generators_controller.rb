class ExcelFileGeneratorsController < ApplicationController

  def new
  end

  def create
    file_data = params[:uploaded_file]
    if file_data.respond_to?(:read)
      
      xml_contents = file_data.read
      doc = Nokogiri::XML(xml_contents)
      doc = doc.remove_namespaces!
      all_the_things = []
      fingerprint   = doc.xpath('//Fingerprint').last['print']
      doc.xpath('//LicenseKey').each_with_index do |file, index|
        licenseKey         = file['id']
        description         = file.xpath("./Description").text
        value         = file.xpath("./value").text
        type         = file.xpath("./Type").text
        
        if (index == 0) 
          all_the_things << [licenseKey, description, value, type,fingerprint]
        else
          all_the_things << [licenseKey, description, value, type]
        end
      end
      excel = Axlsx::Package.new
      book = excel.workbook
      book.add_worksheet(:name => "Basic Worksheet") do |sheet|
        sheet.add_row ["LicenseKey", "Description", "Value", "Type","Fingerprint"]
      all_the_things.each do |data|
        sheet.add_row data
      end
      bold     = { b: true }
      centered = { alignment: { horizontal: :center } }
      sheet.add_style 'A1:E1', bg_color: 'FFF135'
      sheet.add_style 'A1:E1', bold,centered
      sheet.add_style "A1:D#{all_the_things.count + 1}", centered
    end
    excel.serialize("public/xml_data.xlsx")
    redirect_to new_excel_file_generator_path
    else
      logger.error "Bad file_data: #{file_data.class.name}: #{file_data.inspect}"
    end
  end

end
