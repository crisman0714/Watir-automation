
require 'watir'   # the controller
require 'win32ole'

include Watir
#test::unit includes
require 'test/unit' 
require 'test/unit/ui/console/testrunner'
require 'example_logger1'
class TC_MAINTENANCE_REQUESTs < Test::Unit::TestCase
      
  #requires
require 'watir'

#includes
include Watir

ie = IE.new

ie.goto( 'http://localhost:8080/alpha-swipecard-web/faces/pages/index.jsp' )


ie.text_field( :name, 'j_username' ).set( 'cgmanuel' )
ie.button( :name, 'login' ).click

 
ie.frame( :name, 'header' ).link( :id, 'form1:Navigation:navPnl:timesheet' ).click
ie.frame( :name, 'main' ).text_field( :id,'form1:layoutPanel2:button2_clear')
 ie.frame( :name, 'main' ).button( :value, 'Clear' ).click

ie.frame( :name, 'main' ).link( :id, 'form1:timesheetTable:tableRowGroup1:0:tableColumn6:swipeDetail' ).click

ie.frame( :name, 'main' ).link( :id, 'form1:table1:tableRowGroup1:0:tableColumn5:annotateLink' ).click


ie.frame(:name, 'main').text_field(:id, 'form1:taAnnotation_field').set('set')

ie.frame(:name, 'main').text_field(:id, 'form1:taAnnotation').flash
#ie.frame(:name, 'main').text_field(:id, 'form1:taAnnotation').set('hay')
pwede=ie.frame(:name, 'main').text_field(:id, 'form1:taAnnotation_field')
pwede.focus
pwede.set('The quick brown d') 

          
ie.frame( :name, 'main' ).button( :id, 'form1:button1' ).click
end

          
 
 
 
 
 
 
 