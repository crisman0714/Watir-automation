require 'watir'   # the controller
require 'win32ole'
include Watir

#test::unit includes
require 'test/unit' 
require 'test/unit/ui/console/testrunner'

class TC_PIM2_suite < Test::Unit::TestCase
 def startClicker( button , waitTime = 3)

      w = WinClicker.new

      longName = $ie.dir.gsub("/" , "C:/ruby/lib/ruby/site_ruby/1.8/watir/clickJSDialog.rb" )
      shortName = w.getShortFileName(longName)
      c = "start #{shortName }C:/ruby/lib/ruby/site_ruby/1.8/watir/clickJSDialog.rb #{button }
    #{ waitTime} "

      puts "Starting #{c}"

      w.winsystem(c)

      w=nil

    end

def test_a_simplelogin
$ie = IE.new
beginTime = 0
endTime = 0
totalTime=0
test_site = 'http://annatar:8091/webcm/'
puts '## Beginning of test: Enable login'
puts '  '
$ie.goto(test_site)
    $ie.text_field(:name,"username").set("system")
    $ie.text_field(:name,"password").set("system")
    $ie.button(:name,"btnSubmit").click
    $ie.frame(:name,"treeframe").image(:id,"webfx-tree-object-7-plus").click
    $ie.frame(:name,"treeframe").image(:id,"webfx-tree-object-9-image").click
    $ie.frame(:name,"treeframe").link(:id,"webfx-menu-object-90").click
    beginTime = Time.now
    puts beginTime
    $ie.frame(:name,"baseframe").document.all['131'].click
   $ie.frame(:name,"baseframe").wait
   $ie.frame( :name, 'baseframe' ).checkbox( :name, 'relationColMap(1000297)' ).set
    $ie.frame( :name, 'baseframe' ).document.all[ '3854' ].click
    $ie.frame(:name,"baseframe").wait
   
               startClicker("OK" , 3)

            #$ie.button("Submit").click
    endTime = Time.now
    puts endTime
    totalTime= endTime - beginTime
    puts totalTime
 $ie.frame(:name,"baseframe").wait
$ie.frame( :name, 'treeframe' ).link( :id, 'webfx-menu-object-112' ).click
   $ie.frame(:name,"baseframe").wait
$ie.frame( :name, 'baseframe' ).select_list( :name, 'selHistory' ).select( 'Snapshot' )
$ie.frame(:name,"baseframe").wait
$ie.frame( :name, 'baseframe' ).document.all[ '89' ].click   

$ie.frame( :name, 'baseframe' ).checkbox( :name, 'selItemIdList' ).set
$ie.frame( :name, 'baseframe' ).document.all[ '96' ].click #where status is located
$ie.frame( :name, 'baseframe' ).document.all[ '143' ].click # ok button
end
end