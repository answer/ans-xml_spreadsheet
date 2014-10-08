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
  decorator: MyDecorator,
)

class MyDecorator < Draper::Decorator
  include Ans::XmlSpreadsheet::Decorator

  column "name", data_type: "String", width: 100, title: "名前"
  column "age",  data_type: "Number", width: 70,  title: "年齢"
  column "memo", data_type: "Text",   width: 200, title: "備考"
end
```

デコレータは Draper を想定している

セルの内容をデコレーションする必要が無いなら、各パラメータを配列で指定することも可能

```ruby
data.to_xml_spreadsheet(
  columns: [
    ["name", data_type: "String", width: 100, title: "名前"],
    ["age",  data_type: "Number", width: 70,  title: "年齢"],
    ["memo", data_type: "Text",   width: 200, title: "備考"],
  ],
)
```

複数のシートを作成したい場合

```ruby
doc = Ans::XmlSpreadsheet.new
doc.author = "admin"
doc.insert_header = false
doc.push(
  data,
  sheet_name: "sheet1"
  insert_title_row: false,
  decorator: MyDecorator,
  columns: [...],
)
doc.to_xml_spreadsheet
```

## Setting

可能な設定とデフォルト

```ruby
# initializer
Ans::XmlSpreadsheet.configure do |config|
  config.author = "admin"
  config.default_data_type = "String"
  config.default_cell_width = 100
  config.default_insert_title_row = false
end
```

## Decorator

```ruby
class MyDecorator < Draper::Decorator
  include Ans::XmlSpreadsheet::Decorator

  column "name", data_type: "String", width: 100, title: "名前",
    decorates: [:upcase,[:number_to_currency,unit:"USD"]],
    value: ->(name){"name"}

  def upcase(name,value)
    if value.respond_to?(:upcase)
      value.upcase
    else
      value
    end
  end
  def number_to_currency(name,value,opts=nil)
    # raw_value = raw_value(name)
    h.number_to_currency value, opts
  end
end
```

デコレータの `column` メソッドのオプションには、 `data_type` と `width` の他に `decorates` と `value` を指定可能

`decorates` には 「メソッド名、オプション引数」 の配列を指定できる  
オプション引数がなければ項目は配列になってなくても大丈夫  
指定したメソッドが順に カラム名、 整形済みの値 を引数としてコールされる

`value` にはそのカラムの値を指定できる  
`call` メソッドを持つオブジェクトを指定するとアクセスされた時点で一回呼び出されてキャッシュされる

`value` を指定した場合、元の値にアクセスするために、 `raw_value` メソッドが用意されている

### クラスメソッド

```ruby
class MyDecorator < Draper::Decorator
  include Ans::XmlSpreadsheet::Decorator

  column "name", data_type: "String", width: 100, title: "名前"
  column "age",  data_type: "String", width: 100
end

MyDecorator.human_attribute_name("name") #=> "名前"
MyDecorator.human_attribute_name("age")  #=> "age"

MyDecorator.human_attribute_name(0) #=> "名前"
```

`human_attribute_name` で `title` の属性を取得可能

## Contributing

1. Fork it ( https://github.com/answer/ans-xml_spreadsheet/fork )
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create a new Pull Request
