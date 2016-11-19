require 'watir'   # the watir controller
require 'win32ole'
require 'test/unit'
require 'test/unit/ui/console/testrunner'
require 'watir/testUnitAddons'
require 'watir/contrib/enabled_popup' 
require 'watir/winClicker'

include Watir
class TC_MAINTENANCE_REQUEST < Test::Unit::TestCase


 def test_main_proccess
  beginTime = 0
  endTime = 0
  totalTime=0

  $ie = Watir::IE.start("http://manwe:8090/webcm/")
  $ie.text_field(:id, "username").set("system")
  $ie.text_field(:id, "password").set("system")
  $ie.button(:name, "btnSubmit").click
  $ie.frame( :name, 'treeframe' ).image( :id, 'webfx-tree-object-4-plus' ).click
  $ie.frame( :name, 'treeframe' ).link( :id, 'webfx-tree-object-6-anchor' ).click
  $ie.frame( :name, 'treeframe' ).link( :id, 'webfx-menu-object-43' ).click
  $ie.frame( :name, 'baseframe' ).checkbox( :name, 'textSearchAdvancedInd' ).clear
  $ie.frame( :name, 'baseframe' ).text_field( :id, 'wholeTextSearchString' ).set( '
inactive' )
  $ie.frame( :name, 'baseframe' ).button( :name, 'applySearch' ).click
  beginTime = Time.now
$ie.frame(:name,"baseframe").wait
endTime = Time.now
    puts endTime
    totalTime= endTime - beginTime
    puts totalTime 
    puts "ellapse time"
$ie.close
end
end