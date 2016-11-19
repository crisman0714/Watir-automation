require 'watir'   # the controller
include Watir

#test::unit includes
require 'test/unit' 
require 'test/unit/ui/console/testrunner'
require 'watir/contrib/ie-new-process'
class TC_enable_suite < Test::Unit::TestCase

def test_a_simplelogin
 ies = []
3.times do
 # ie = Watir::IE.new_process
  
  test_site = 'http://palantiri/enable51/enable.login.web/login.aspx'
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
  
     $ie.close
  
  ies<<$ie
end

end 
#ef test_a_simplelogin2

#test_site = 'http:/palantiri/enable51/enable.login.web/login.aspx'
#puts '## Beginning of test: Enable login'
#puts '  '
#$ie = IE.new
#$ie.goto(test_site)

#puts ' enter "Login: in the text field'
#$ie.text_field(:name, "txtUserLogin").set("admin1")


#puts ' enter "Password: in the text field'
#$ie.text_field(:name, "txtUserPwd").set("om")

#puts 'Step 3: click the submit button'
 #  $ie.button(:name, "btnSubmit").click

#end 
end