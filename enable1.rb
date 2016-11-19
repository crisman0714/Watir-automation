require 'watir'   # the controller
include Watir

#test::unit includes
require 'test/unit' 
require 'test/unit/ui/console/testrunner'

class TC_enable_suite < Test::Unit::TestCase

def test_a_simplelogin

test_site = 'http://eonas/enable51/enable.login.web/login.aspx'
puts '## Beginning of test: Enable login'
puts '  '
$ie = IE.new
$ie.goto(test_site)

puts ' enter "Login: in the text field'
$ie.text_field(:name, "txtUserLogin").set("admin")


puts ' enter "Login: in the text field'
$ie.text_field(:name, "txtUserPwd").set("om")

puts 'Step 3: click the submit button'
   $ie.button(:name, "btnSubmit").click

test_site2 = 'http://eonas/enable51/enable.dam.web/index.aspx'

$ie.goto(test_site2)
 #$ie.caption(:name, "all items").click
$ie.document.all['200'].click
 $ie.text_field(:id, "txtJumptoName").set("samurai")
 $ie.button(:id, "command-button").click
end

end