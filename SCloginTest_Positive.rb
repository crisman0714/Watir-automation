require 'watir'   # the controller
require 'win32ole'

include Watir
#test::unit includes
require 'test/unit' 
require 'test/unit/ui/console/testrunner'
require 'example_logger1'
class TC_MAINTENANCE_REQUESTs < Test::Unit::TestCase
      
   def division_Text_Logger
     
    $ie = IE.new
    filePrefix = "SCLogin_logger"
    $logger = LoggerFactory.start_xml_logger(filePrefix) 
    $ie.set_logger($logger)
    end
   def test_SC_Login
       division_Text_Logger
       test_site = 'http://localhost:8080/alpha-swipecard-web'
        $ie.goto(test_site)
        excel = WIN32OLE::new("excel.Application")
        workbook = excel.Workbooks.Open("c:\\SClogin.xls") # directory Path where the test data is located
        worksheet = workbook.WorkSheets(2)
        worksheet.Select
        line = '2'

        while
            test_id= worksheet.Range("a#{line}")["Value"]  
            requestor_username= worksheet.Range("b#{line}")["Value"]  
            requestor_password= worksheet.Range("c#{line}")["Value"]  
            testValidator = worksheet.Range("d#{line}")["Value"]
            description = worksheet.Range("e#{line}")["Value"]
            testCaseId = worksheet.Range("f#{line}")["Value"]
            
            $logger.log(" ")
            $logger.log(test_id)
          
            $ie.text_field(:name,"j_username").set(requestor_username)
            $logger.log("Action: Entered " + requestor_username + " as username")
            $ie.text_field(:name,"j_password").set(requestor_password)
            $logger.log("Action: Entered " + requestor_password + " as password")
            $ie.button(:value, "Log In").click
            $logger.log( "Action: Clicked Login Button")
          
        if($ie.frame( :name, 'main').contains_text( testValidator ))
          #if($ie.pageContainsText(testValidator))
          $logger.log("Passed")
            $logger.log("Test Description:" +testCaseId+":"+ description)
          else 
            $logger.log("Failed") 
            $logger.log("Test Description:" +testCaseId+":"+ description)
            
          end
             $ie.goto(test_site)
             line.succ!
               
             end
            $ie.frame( :name, 'header' ).link( :id, 'form1:Header:layoutPanel2:layoutPanel1:logout' ).click 
          $ie.close
    end
 end
          
 
      