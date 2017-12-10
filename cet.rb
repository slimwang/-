# encoding: UTF-8
require 'rubygems'
require 'mechanize'
require 'roo-xls'
require 'simple-xls'

#==================初始化数据==========================

xls = SimpleXLS.new ['4/6','Listening','Reading','Writing','Total','College','Name']

#打开文件
excel = Roo::Excel.new("id.xls")
excel.default_sheet = excel.sheets.first
last_row = excel.last_row

#读入数组
id       = [] #准考证号
name     = [] #姓名

1.upto(last_row) do |line|
  id       << excel.cell(line,'B')
  name     << excel.cell(line,'C')
end
id.shift    #除去header
name.shift  #除去header

#=====================查询成绩===========================

1.upto(last_row-1) do |r|
  agent = Mechanize.new
  page = agent.get('http://cet.99sushe.com/')

  #构造表单
  search_form = page.form('searchform')
  search_form["id"] = id[r]
  search_form["name"] = name[r]

  #提交表单
  page = agent.submit(search_form, search_form.buttons.first)

  #数据处理
  score = page.body.split(',')
  xls.push score
end

#储存为excel文件
File.open('output.xls', 'w+') { |f| f.puts xls }
