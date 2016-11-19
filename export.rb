require 'watir'   # the watir controller
require 'win32ole'
require 'test/unit'
require 'test/unit/ui/console/testrunner'
require 'watir/testUnitAddons'
require 'watir/contrib/enabled_popup' 
require 'watir/winClicker'

include Watir
class TC_MAINTENANCE_REQUEST < Test::Unit::TestCase
def complete_func
    $ie.frame(:name,"baseframe").wait
    #onse=$ie.frame( :name, 'baseframe' ).contains_text("Processing")
    while 
    if $ie.frame( :name, 'baseframe' ).document.all['102'].innerText.match('Queued')
            puts "Still on Process"
        $ie.frame( :name, 'baseframe' ).button(:value, "Refresh").click
        sleep 1
        else 
          if $ie.frame( :name, 'baseframe' ).document.all['102'].innerText.match('Processing')
           puts "Still on Process"
        $ie.frame( :name, 'baseframe' ).button(:value, "Refresh").click
        sleep 1
        else
  $ie.frame( :name, 'baseframe' ).button(:value, "Refresh").click
        puts"Ok Na"
      end
     end
    end
  end

 def test_main_proccess
  beginTime = 0
  endTime = 0
  totalTime=0

  $ie = Watir::IE.start("http://manwe:8090/webcm/")
  $ie.text_field(:id, "username").set("system")
  $ie.text_field(:id, "password").set("system")
  $ie.button(:name, "btnSubmit").click
  $ie.frame( :name, 'treeframe' ).image( :id, 'webfx-tree-object-4-plus' ).click
  $ie.frame( :name, 'treeframe' ).link( :id, 'webfx-tree-object-5-anchor' ).click
  $ie.frame( :name, 'treeframe' ).link( :id, 'webfx-menu-object-37' ).click
  $ie.frame( :name, 'baseframe' ).document.all[ '92' ].click
beginTime = Time.now
  $ie.frame(:name,"baseframe").wait

complete_func
 
 endTime = Time.now
    puts endTime
    totalTime= endTime - beginTime
    puts totalTime 
    puts "ellapse time"
  $ie.close
end
end