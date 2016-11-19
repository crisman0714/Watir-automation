require 'watir'   # the controller
include Watir

#test::unit includes
require 'test/unit' 
require 'test/unit/ui/console/testrunner'

class TC_PIM2_suite < Test::Unit::TestCase

def test_a_simplelogin

test_site = 'http://baggins:8080/pim2/login.action'
puts '## Beginning of test: Enable login'
puts '  '
$ie = IE.new
$ie.goto(test_site)

puts ' enter "Login: in the text field'
$ie.text_field(:name, "userForm.userName").set("druser")


puts ' enter "Login: in the text field'
$ie.text_field(:name, "userForm.userPassword").set("password")

puts 'Step 3: click the submit button'
   $ie.button(:value, "Login").click
  
 $ie.image(:index,2).click
 # test_site2 = 'http://gondor:8080/pim2/searchHome.action'

#$ie.goto(test_site2)
$ie.show_frames
  $ie.show_images
$ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.searchKeywords").set("cashew")
$ie.frame("main").frame(:name,"query").button(:name,"method:executeSearchQuery").click




end 

end