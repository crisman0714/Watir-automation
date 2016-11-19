
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
   def test_SC6_Login
       division_Text_Logger
       test_site = 'http://localhost:8080/alpha-swipecard-web'
        $ie.goto(test_site)
        excel = WIN32OLE::new("excel.Application")
        workbook = excel.Workbooks.Open("c:\\SClogin.xls") # directory Path where the test data is located
        worksheet = workbook.WorkSheets(3)
        worksheet.Select
        line = '2'

#Login into the system
        $ie.text_field( :name, 'j_username' ).set( 'jfcapule' )
        $ie.text_field( :name, 'j_password' ).set( 'joecapule167' )
        $ie.button( :name, 'login' ).click
 ######################################################################           
  #shift in out Codes
        $ie.frame( :name, 'header' ).link( :id, 'form1:Navigation:navPnl:shift' ).click
        $logger.log("Shift in out") 
 while
            test_id= worksheet.Range("a#{line}")["Value"]  
            dateFromRange= worksheet.Range("b#{line}")["Value"]  
            dateToRange= worksheet.Range("c#{line}")["Value"]  
            testValidator = worksheet.Range("d#{line}")["Value"]
            description = worksheet.Range("e#{line}")["Value"]
            testCaseId = worksheet.Range("f#{line}")["Value"]
            testCaseNumber= worksheet.Range("g#{line}")["Value"]
           
          
            
            
            $logger.log(" ")
            $logger.log(test_id)
 
          $ie.frame( :name, 'main' ).text_field( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal_field' ).set(dateFromRange)
          $logger.log("Action: Entered " + dateFromRange + " as Date From")
          $ie.frame( :name, 'main' ).text_field( :id, 'form1:layoutPanel2:groupPanel1:dateToCal_field' ).set(dateToRange)
          $logger.log("Action: Entered " + dateToRange + " as Date To")
          $ie.frame( :name, 'main' ).button( :id, 'form1:layoutPanel2:searchBtn' ).click
          $ie.frame( :name, 'main' ).button( :id, 'form1:layoutPanel2:searchBtn' ).click
  if($ie.frame( :name, 'main').contains_text( testValidator ))
          worksheet.Range("h#{line}").Value="Pass"
          worksheet.range("h#{line}").Interior['ColorIndex'] =50
          $logger.log("Pass")
          $logger.log("Test Description:"+testCaseNumber+":" +testCaseId+":"+ description)
          else 
             worksheet.Range("h#{line}").Value="Fail"
            worksheet.range("h#{line}").Interior['ColorIndex'] =27
           $logger.log("Fail") 
              $logger.log("Test Description:"+testCaseNumber+":" +testCaseId+":"+ description)
            
          end
             
             line.succ!
                
             end
             
  ######################################################################           
  #Timesheet Correlation Codes
              #line = '2'
     #       excel = WIN32OLE::new("excel.Application")
       # workbook = excel.Workbooks.Open("c:\\SClogin.xls") # directory Path where the test data is located
       # worksheet = workbook.WorkSheets(3)
       # worksheet.Select
        line = '23'
            $ie.frame( :name, 'header' ).link( :id, 'form1:Navigation:navPnl:timesheet' ).click
            $logger.log("Timesheet") 
                        
 while
            test_id= worksheet.Range("a#{line}")["Value"]  
            dateFromRange= worksheet.Range("b#{line}")["Value"]  
            dateToRange= worksheet.Range("c#{line}")["Value"]  
            testValidator = worksheet.Range("d#{line}")["Value"]
            description = worksheet.Range("e#{line}")["Value"]
            testCaseId = worksheet.Range("f#{line}")["Value"]
            testCaseNumber= worksheet.Range("g#{line}")["Value"]
             
            
            $logger.log(" ")
            $logger.log(test_id)
 
          $ie.frame( :name, 'main' ).text_field( :id, 'form1:SingleUserSearchParameters:layoutPanel1:dateGroupPanel:dateFromCal_field' ).set(dateFromRange)
          $logger.log("Action: Entered " + dateFromRange + " as Date From")
          $ie.frame( :name, 'main' ).text_field( :id,  'form1:SingleUserSearchParameters:layoutPanel1:dateGroupPanel:dateToCal_field' ).set(dateToRange)
          $logger.log("Action: Entered " + dateToRange + " as Date To")
          $ie.frame( :name, 'main' ).button( :id, 'form1:SingleUserSearchParameters:layoutPanel1:searchButton'  ).click
          $ie.wait
           $ie.frame( :name, 'main' ).button( :id, 'form1:SingleUserSearchParameters:layoutPanel1:searchButton'  ).click
           $ie.wait
           $ie.frame( :name, 'main' ).button( :id, 'form1:SingleUserSearchParameters:layoutPanel1:searchButton'  ).click
  if($ie.frame( :name, 'main').contains_text( testValidator ))
          worksheet.Range("h#{line}").Value="Pass"
            worksheet.range("h#{line}").Interior['ColorIndex'] =50
          $logger.log("Pass")
              $logger.log("Test Description:"+testCaseNumber+":" +testCaseId+":"+ description)
          else
             worksheet.Range("h#{line}").Value="Fail"
            worksheet.range("h#{line}").Interior['ColorIndex'] =27
           $logger.log("Fail") 
            $logger.log("Test Description:"+testCaseNumber+":" +testCaseId+":"+ description)
            
          end
             line.succ!
               
             end
             $ie.frame( :name, 'header' ).link( :id, 'form1:Header:layoutPanel2:layoutPanel1:logout' ).click 
          $ie.close
          workbook.save
           workbook.close
           excel.Quit   
    end
 end
 
 
 
 
 
 
 
 
 
 