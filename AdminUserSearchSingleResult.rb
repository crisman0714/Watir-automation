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
   def test_SC2_Login
       division_Text_Logger
       test_site = 'http://localhost:8080/alpha-swipecard-web'
        $ie.goto(test_site)
        excel = WIN32OLE::new("excel.Application")
        workbook = excel.Workbooks.Open("c:\\SClogin.xls") # directory Path where the test data is located
        worksheet = workbook.WorkSheets(7)
        worksheet.Select
        line = '2'

#Login into the system
#$ie.text_field( :name, 'j_username' ).set( 'cgmanuel' )
  #    $ie.text_field( :name, 'j_password' ).set( 'kilouwa31' )
    $ie.text_field( :name, 'j_username' ).set( 'crcortez' )
    $ie.text_field( :name, 'j_password' ).set( 'onesecret' )
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
            worksheet.range("h#{line}").Interior['ColorIndex'] =28        
            $logger.log("Fail") 
              $logger.log("Test Description:"+testCaseNumber+":" +testCaseId+":"+ description)
            
          end
             
             line.succ!
                
             end
          
          
                ######################################################################           
  #Timesheet Correlation Codes
        line = '21'   
        $ie.frame( :name, 'header' ).link( :id, 'form1:Navigation:navPnl:timesheet' ).click
        $logger.log("Timesheet Correlation") 
 while
            test_id= worksheet.Range("a#{line}")["Value"]  
            dateFromRange= worksheet.Range("b#{line}")["Value"]  
            dateToRange= worksheet.Range("c#{line}")["Value"] 
            testName = worksheet.Range("d#{line}")["Value"]
            testValidator1 = worksheet.Range("e#{line}")["Value"]
            description = worksheet.Range("f#{line}")["Value"]
            testCaseId = worksheet.Range("g#{line}")["Value"]
            testCaseNumber= worksheet.Range("h#{line}")["Value"]
            
            
            $logger.log(" ")
            $logger.log(test_id)
 
          $ie.frame( :name, 'main' ).text_field( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal_field' ).set(dateFromRange)
          $logger.log("Action: Entered " + dateFromRange + " as Date From")
          $ie.frame( :name, 'main' ).text_field( :id, 'form1:layoutPanel2:groupPanel1:dateToCal_field' ).set(dateToRange)
          $logger.log("Action: Entered " + dateToRange + " as Date To")
          $ie.frame( :name, 'main' ).text_field( :id, 'form1:layoutPanel2:groupPanel2:nameTxt_field' ).set(testName)
           $logger.log("Action: Entered " + testName + " as Name")
          $ie.frame( :name, 'main' ).button( :id, 'form1:layoutPanel2:searchBtn' ).click
          $ie.wait
          $ie.frame( :name, 'main' ).button( :id, 'form1:layoutPanel2:searchBtn' ).click
          $ie.wait
          $ie.frame( :name, 'main' ).button( :id, 'form1:layoutPanel2:searchBtn' ).click
  if($ie.frame( :name, 'main').contains_text( testValidator1 ))
          worksheet.Range("i#{line}").Value="Pass"
          worksheet.range("i#{line}").Interior['ColorIndex'] =50
          $logger.log("Pass")
          $logger.log(testCaseNumber+":" +testCaseId+":"+ description)
          else 
            worksheet.Range("i#{line}").Value="Fail"
            worksheet.range("i#{line}").Interior['ColorIndex'] =28        
           $logger.log("Fail") 
           $logger.log(testCaseNumber+":" +testCaseId+":"+ description)
            
          end
             #$ie.goto(test_site)
             line.succ!
               
             end
   ######################################################################           
  #Audit Trail Codes
        line = '46'   
             
        $ie.frame( :name, 'header' ).link( :id, 'form1:Navigation:navPnl:audit').click
        $logger.log("Audit Trail") 
 while
            test_id= worksheet.Range("a#{line}")["Value"]  
            dateFromRange= worksheet.Range("b#{line}")["Value"]  
            dateToRange= worksheet.Range("c#{line}")["Value"] 
            testName = worksheet.Range("d#{line}")["Value"]
            testValidator1 = worksheet.Range("e#{line}")["Value"]
            description = worksheet.Range("f#{line}")["Value"]
            testCaseId = worksheet.Range("g#{line}")["Value"]
            testCaseNumber = worksheet.Range("h#{line}")["Value"]
            
            $logger.log(" ")
            $logger.log(test_id)
 
          $ie.frame( :name, 'main' ).text_field( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal_field' ).set(dateFromRange)
          $logger.log("Action: Entered " + dateFromRange + " as Date From")
          $ie.frame( :name, 'main' ).text_field( :id, 'form1:layoutPanel2:groupPanel1:dateToCal_field' ).set(dateToRange)
          $logger.log("Action: Entered " + dateToRange + " as Date To")
          $ie.frame( :name, 'main' ).select_list( :id, 'form1:layoutPanel2:reportTypeDropDown_list' ).select(testName)
           $logger.log("Action: Entered " + testName + " as Unname field")
          $ie.frame( :name, 'main' ).button( :id, 'form1:layoutPanel2:searchBtn' ).click
  $ie.wait
  if($ie.frame( :name, 'main' ).contains_text( testValidator1 )) 
           worksheet.Range("i#{line}").Value="Pass"
            worksheet.range("i#{line}").Interior['ColorIndex'] =50
          $logger.log("Pass")
          $logger.log("Test Description:" + testCaseNumber + ":" + testCaseId + ":" + description)
          else 
             worksheet.Range("i#{line}").Value="Fail"
            worksheet.range("i#{line}").Interior['ColorIndex'] =28         
            $logger.log("Fail") 
             $logger.log(testCaseNumber+":" +testCaseId+":"+ description)
          end
             #$ie.goto(test_site)
             line.succ!
               
             end
            $ie.frame( :name, 'header' ).link( :id, 'form1:Header:layoutPanel2:layoutPanel1:logout' ).click 
          $ie.close
           workbook.save
           #workbook.close
           excel.Quit   
    end
 end
 
 