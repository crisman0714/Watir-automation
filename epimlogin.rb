require 'watir'   # the controller
require 'win32ole'
include Watir

#test::unit includes
require 'test/unit' 
require 'test/unit/ui/console/testrunner'

class TC_PIM2_suite < Test::Unit::TestCase
 def test_a_simplelogin
 
excel = WIN32OLE::new("excel.Application")
workbook = excel.Workbooks.Open("c:\\example\\example_test_epim_login1.xls")

worksheet = workbook.WorkSheets(1) # get first workbook
worksheet.Select    # Just to make sure macros are executed, if you sheet doesn't have macros you can skip this step.
$ie = IE.new
line = '1'
while 
 value2 = worksheet.Range("b#{line}")["Value"]  

value = worksheet.Range("a#{line}")["Value"]      # Gets the value at cell a12 in worksheet.


test_site = 'http://doriath:8092/webcm/'
puts '## Beginning of test: Enable login'
puts '  '

$ie.goto(test_site)
    $ie.text_field(:name,"username").set(value)
    $ie.text_field(:name,"password").set(value2)
    $ie.button(:name,"btnSubmit").click
    #$ie.frame(:name,"treeframe").image(:id,"webfx-tree-object-5-plus").click
  #$ie.frame(:name,"baseframe").wait

   
line.succ!
 end
end
end