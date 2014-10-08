require "ans/xml_spreadsheet/version"

module Ans
  module XmlSpreadsheet
    autoload "Document", "ans/xml_spreadsheet/document.rb"

    include ActiveSupport::Configurable

    configure do |config|
      config.default_author = "admin"
      config.default_data_type = "String"
      config.default_cell_width = 100
      config.default_insert_title_row = false
      config.default_sheet_name_generator = ->(index){"sheet#{index+1}"}
    end

    def self.add_ms_excel_mime_type!(format: :xls)
      Mime::Type.register "application/vnd.ms-excel", format
    end

    module Generatable
      def self.included(m)
        m.send :include, InstanceMethods
      end

      module InstanceMethods
        def to_xml_spreadsheet(sheet_name: nil, insert_title_row: false, columns: nil)
          doc = Ans::XmlSpreadsheet::Document.new
          doc.push(
            self,
            sheet_name: sheet_name,
            insert_title_row: insert_title_row,
            columns: columns,
          )
          doc.to_xml_spreadsheet
        end
      end
    end
  end
end

if defined?(Array)
  class Array
    include Ans::XmlSpreadsheet::Generatable
  end
end
if defined?(ActiveRecord::Base)
  class ActiveRecord::Base
    include Ans::XmlSpreadsheet::Generatable
  end
end
