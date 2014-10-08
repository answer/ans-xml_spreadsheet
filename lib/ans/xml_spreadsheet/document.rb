require "rexml/document"

class Ans::XmlSpreadsheet::Document
  attr_accessor *%i{
    author
    default_data_type
    default_width
    default_insert_title_row
    sheet_name_generator
    created
  }

  [
    [:author,:default_author],
    :default_data_type,
    :default_width,
    :default_insert_title_row,
    [:sheet_name_generator,:default_sheet_name_generator],
  ].each do |method,default_method|
    default_method ||= method
    define_method method do
      instance_variable_get(:"@#{method}") || Ans::XmlSpreadsheet.config.send(default_method)
    end
  end

  def push(data, sheet_name: nil, insert_title_row: false, columns: nil)
    unless columns
      columns = []
      row = data.first
      case row
      when ActiveRecord::Base
        row.attribute_names.each do |name|
          columns << [name]
        end
      when Array
        row.each_with_index do |value,i|
          columns << [i]
        end
      end
    else
      columns = columns.each_with_index.map do |(name,opts),i|
        unless opts
          case name
          when Hash
            opts = name
            name = i
          end
        end
        [name,opts]
      end
    end

    row_count = data.size
    if insert_title_row
      row_count += 1
    end

    sheets.push(
      data: data,
      sheet_name: sheet_name || sheet_name_generator.call(sheets.size),
      insert_title_row: insert_title_row,
      columns: columns,
      row_count: row_count,
    )
  end

  def to_xml_spreadsheet
    doc = REXML::Document.new
    doc << REXML::XMLDecl.new('1.0', 'UTF-8')
    doc << REXML::Instruction.new("mso-application", %{progid="Excel.Sheet"})

    work_book = doc.add_element(
      "Workbook",
      "xmlns" => "urn:schemas-microsoft-com:office:spreadsheet",
      "xmlns:o" => "urn:schemas-microsoft-com:office:office",
      "xmlns:x" => "urn:schemas-microsoft-com:office:excel",
      "xmlns:ss" => "urn:schemas-microsoft-com:office:spreadsheet",
      "xmlns:html" => "http://www.w3.org/TR/REC-html40",
    )
    document_properties = work_book.add_element(
      "DocumentProperties",
      "xmlns" => "urn:schemas-microsoft-com:office:office",
    )
    document_properties.add_element("Author").add_text(author)
    document_properties.add_element("Created").add_text((created || Time.now.gmtime).iso8601)

    excel_work_book = work_book.add_element(
      "ExcelWorkbook",
      "xmlns" => "urn:schemas-microsoft-com:office:excel",
    )
    excel_work_book.add_element("ProtectStructure").add_text("False")
    excel_work_book.add_element("ProtectWindows").add_text("False")

    styles = work_book.add_element("Styles")
    style = styles.add_element(
      "Style",
      "ss:ID" => "Default",
      "ss:Name" => "Normal",
    )
    style.add_element("Alignment","ss:Vertical"=>"Center")
    style.add_element("Borders")
    style.add_element("Interior")
    style.add_element("NumberFormat")
    style.add_element("Protection")

    style = styles.add_element("Style", "ss:ID" => "sText")
    style.add_element(
      "Alignment",
      "ss:Horizontal" => "Left",
      "ss:Vertical" => "Top",
      "ss:WrapText" => "1",
    )

    style = styles.add_element("Style", "ss:ID" => "sNumber")
    style.add_element(
      "Alignment",
      "ss:Vertical" => "Top",
    )

    style = styles.add_element("Style", "ss:ID" => "sString")
    style.add_element(
      "Alignment",
      "ss:Horizontal" => "Left",
      "ss:Vertical" => "Top",
    )

    sheets.each do |sheet|
      work_sheet = work_book.add_element("Worksheet", "ss:Name" => sheet[:sheet_name])
      table = work_sheet.add_element(
        "Table",
        "ss:ExpandedColumnCount" => sheet[:columns].size,
        "ss:ExpandedRowCount" => sheet[:row_count],
      )
      sheet[:columns].each do |column,opts|
        data_type = opts[:data_type] || default_data_type

        attrs = {
          "ss:StyleID" => "s#{data_type}",
          "ss:Width" => opts[:width] || default_width,
        }

        case data_type
        when "Number"
          attrs["ss:AutoFitWidth"] = 1
        end

        table.add_element("Column",attrs)
      end
      if sheet[:insert_title_row]
        row = table.add_element("Row","ss:AutoFitHeight" => 0)
        sheet[:columns].each do |column,opts|
          human_attribute_name = opts[:title] || column

          row.add_element("Cell")
            .add_element("Data","ss:Type" => "String")
            .add_text(human_attribute_name)
        end
      end
      sheet[:data].each do |data_row|
        case data_row
        when ActiveRecord::Base
          data = data_row.attributes
        when Array
          data = data_row
        end

        if data
          row = table.add_element("Row","ss:AutoFitHeight" => 1)
          sheet[:columns].each_with_index do |(column,opts),i|
            data_type = opts[:data_type] || default_data_type
            if data_type == "Text"
              data_type = "String"
            end

            row.add_element("Cell")
              .add_element("Data","ss:Type" => data_type)
              .add_text(data[column])
          end
        end
      end
    end

    doc
  end

  private

  def sheets
    @sheets ||= []
  end
end
