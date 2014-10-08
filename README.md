# Ans::XmlSpreadsheet

CSV、配列、ActiveRecord などのコレクションから xml spreadsheet を生成する

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'ans-xml_spreadsheet'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install ans-xml_spreadsheet

## Usage

```ruby
data = CSV.read("data.csv")
data.to_xml_spreadsheet

data = Model.awesome_scope
data.to_xml_spreadsheet
```

戻り値の型は `REXML::Document`

デフォルトはすべてのデータは String 型、幅 100 のセルとして整形される

`Array`, `ActiveRecord::Base` に対して、 `to_xml_spreadsheet` メソッドが生える

`Ans::XmlSpreadsheet::Generatable` を include すると `to_xml_spreadsheet` が生える

### Options

```ruby
data.to_xml_spreadsheet(
  author: "admin",
  sheet_name: "sheet1",
  insert_title_row: false,
  columns: [
    ["name", data_type: "String", width: 100, title: "名前"],
    ["age",  data_type: "Number", width: 70,  title: "年齢"],
    ["memo", data_type: "Text",   width: 200, title: "備考"],
  ],
)
```

`insert_title_row` に true を指定すると、もともとのデータに加えて、カラムの名前の行を追加する

`columns` には、列情報を指定する

* `data_type` : String, Number, Text を使用可能
* `width` : セルの幅
* `title` : `insert_title_row` が true の場合に追加される行に使用されるカラムの名前

複数のシートを作成したい場合

```ruby
doc = Ans::XmlSpreadsheet::Document.new
doc.author = "admin"
doc.default_data_type = "String"
doc.default_width = 100
doc.default_insert_title_row = false
doc.sheet_name_generator = ->(index){"sheet#{index+1}"}

doc.push(
  data,
  sheet_name: "sheet1"
  insert_title_row: false,
  columns: [...],
)
doc.push(
  ...
)

doc.to_xml_spreadsheet
```

## Setting

可能な設定とデフォルト

```ruby
# initializer
Ans::XmlSpreadsheet.configure do |config|
  config.default_author = "admin"
  config.default_data_type = "String"
  config.default_width = 100
  config.default_insert_title_row = false
  config.sheet_name_generator = ->(index){"sheet#{index+1}"}
end

#Ans::XmlSpreadsheet.add_ms_excel_mime_type!(format: :xls)
```

`add_ms_excel_mime_type!` をコールすると、 `Mime::Type` に `application/vnd.ms-excel` が追加される  
オプションで `format` を渡せる  
省略すると `xls`

## Contributing

1. Fork it ( https://github.com/answer/ans-xml_spreadsheet/fork )
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create a new Pull Request
