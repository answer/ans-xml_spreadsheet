# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'ans/xml_spreadsheet/version'

Gem::Specification.new do |spec|
  spec.name          = "ans-xml_spreadsheet"
  spec.version       = Ans::XmlSpreadsheet::VERSION
  spec.authors       = ["sakai shunsuke"]
  spec.email         = ["sakai@ans-web.co.jp"]
  spec.summary       = %q{コレクションから xml spreadsheet を生成する}
  spec.description   = %q{CSV、配列、ActiveRecord などのコレクションから xml spreadsheet を生成する}
  spec.homepage      = "https://github.com/answer/ans-xml_spreadsheet"
  spec.license       = "MIT"

  spec.files         = `git ls-files -z`.split("\x0")
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.add_development_dependency "bundler", "~> 1.7"
  spec.add_development_dependency "rake", "~> 10.0"
end
