
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
   def test_SC3_Login
       division_Text_Logger
       test_site = 'http://localhost:8080/alpha-swipecard-web'
        $ie.goto(test_site)
        excel = WIN32OLE::new("excel.Application")
        workbook = excel.Workbooks.Open("c:\\SClogin.xls") # directory Path where the test data is located
        worksheet = workbook.WorkSheets(6)
        worksheet.Select
        line = '2'

#Login into the system
        $ie.text_field( :name, 'j_username' ).set( 'cgmanuel' )
        $ie.text_field( :name, 'j_password' ).set( 'kilouwa31' )
        $ie.button( :name, 'login' ).click

        $ie.frame( :name, 'header' ).link( :id, 'form1:Navigation:navPnl:shift' ).click
        $logger.log("Shift in out") 
 while
            test_id= worksheet.Range("a#{line}")["Value"]  
            dateFromRange= worksheet.Range("b#{line}")["Value"]  
            dateToRange= worksheet.Range("c#{line}")["Value"] 
            testName = worksheet.Range("d#{line}")["Value"]
            testValidator1 = worksheet.Range("e#{line}")["Value"]
            testValidator2 = worksheet.Range("f#{line}")["Value"]
            description = worksheet.Range("g#{line}")["Value"]
            testCaseId = worksheet.Range("h#{line}")["Value"]
            testCaseNumber = worksheet.Range("i#{line}")["Value"]
            
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
  if($ie.frame( :name, 'main').contains_text( testValidator1 ) && $ie.frame( :name, 'main').contains_text( testValidator2))
          worksheet.Range("j#{line}").Value="Pass"
          worksheet.range("j#{line}").Interior['ColorIndex'] =50
          $logger.log("Pass")
          $logger.log("Test Description:"+testCaseNumber+":" +testCaseId+":"+ description)
          else 
           worksheet.Range("j#{line}").Value="Fail"
            worksheet.range("j#{line}").Interior['ColorIndex'] =27
           $logger.log("Fail") 
            $logger.log("Test Description:"+testCaseNumber+":" +testCaseId+":"+ description)
            
          end
             #$ie.goto(test_site)
             line.succ!
               
             end
  ######################################################################           
  #Timesheet Correlation Codes
        line = '16'   
             
        $ie.frame( :name, 'header' ).link( :id, 'form1:Navigation:navPnl:timesheet').click
        $logger.log("Timesheet Correlation") 
 while
            test_id= worksheet.Range("a#{line}")["Value"]  
            dateFromRange= worksheet.Range("b#{line}")["Value"]  
            dateToRange= worksheet.Range("c#{line}")["Value"] 
            testName = worksheet.Range("d#{line}")["Value"]
            testValidator1 = worksheet.Range("e#{line}")["Value"]
            testValidator2 = worksheet.Range("f#{line}")["Value"]
            description = worksheet.Range("g#{line}")["Value"]
            testCaseId = worksheet.Range("h#{line}")["Value"]
            testCaseNumber = worksheet.Range("i#{line}")["Value"]
            
            $logger.log(" ")
            $logger.log(test_id)
 
          $ie.frame( :name, 'main' ).text_field( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal_field' ).set(dateFromRange)
          $logger.log("Action: Entered " + dateFromRange + " as Date From")
          $ie.frame( :name, 'main' ).text_field( :id, 'form1:layoutPanel2:groupPanel1:dateToCal_field' ).set(dateToRange)
          $logger.log("Action: Entered " + dateToRange + " as Date To")
          $ie.frame( :name, 'main' ).text_field( :id, 'form1:layoutPanel2:groupPanel2:nameTxt_field' ).set(testName)
           $logger.log("Action: Entered " + testName + " as Name")
          $ie.frame( :name, 'main' ).button( :id, 'form1:layoutPanel2:searchBtn' ).click
          $ie.frame( :name, 'main' ).button( :id, 'form1:layoutPanel2:searchBtn' ).click
  if($ie.frame( :name, 'main').contains_text( testValidator1 ) && $ie.frame( :name, 'main').contains_text( testValidator2))
          worksheet.Range("j#{line}").Value="Pass"
          worksheet.range("j#{line}").Interior['ColorIndex'] =50
          $logger.log("Pass")
          $logger.log("Test Description:"+testCaseNumber+":" +testCaseId+":"+ description)
          else 
           worksheet.Range("j#{line}").Value="Fail"
            worksheet.range("j#{line}").Interior['ColorIndex'] =27
           $logger.log("Fail") 
            $logger.log("Test Description:"+testCaseNumber+":" +testCaseId+":"+ description)
            
          end
             #$ie.goto(test_site)
             line.succ!
               
             end


######################################################################           
  #Who's here Codes
        line = '29'   
             
        $ie.frame( :name, 'header' ).link( :id, 'form1:Navigation:navPnl:whoIsHere').click
        $logger.log("Who's Here") 
 while
            test_id= worksheet.Range("a#{line}")["Value"]  
            date= worksheet.Range("b#{line}")["Value"]  
            timeHour= worksheet.Range("c#{line}")["Value"] 
            timeMins = worksheet.Range("d#{line}")["Value"]
            timePmAm = worksheet.Range("e#{line}")["Value"]
            selector = worksheet.Range("f#{line}")["Value"]
            testValidator1 = worksheet.Range("g#{line}")["Value"]
            testValidator2 = worksheet.Range("h#{line}")["Value"]
            description = worksheet.Range("i#{line}")["Value"]
            testCaseId = worksheet.Range("j#{line}")["Value"]
            testCaseNumber = worksheet.Range("k#{line}")["Value"]
            
            $logger.log(" ")
            $logger.log(test_id)
 
          $ie.frame( :name, 'main' ).text_field( :id, 'whosHereForm:datePicker_field' ).set(date)
          $logger.log("Action: Entered " + date+ " as Date ")
          $ie.frame( :name, 'main' ).text_field( :id, 'whosHereForm:hour_field' ).set( timeHour)
          $logger.log("Action: Entered " + timeHour + " as Hour")
          $ie.frame( :name, 'main' ).text_field( :id, 'whosHereForm:min_field' ).set(timeMins)
          $logger.log("Action: Entered " + timeMins + " as Minutes")
          $ie.frame( :name, 'main' ).select_list( :id, 'whosHereForm:meridiem_list' ).select(timePmAm )
          $logger.log("Action: Entered " + timePmAm  + " as PM/AM")
          $ie.frame( :name, 'main' ).select_list( :id, 'whosHereForm:reportType_list' ).select(selector)
           $logger.log("Action: Entered " + selector + " as Selector for report")
          $ie.frame( :name, 'main' ).button( :id, 'whosHereForm:searchBtn' ).click
 
  if($ie.frame( :name, 'main').contains_text( testValidator1 ) && $ie.frame( :name, 'main').contains_text( testValidator2))
          worksheet.Range("l#{line}").Value="Pass"
          worksheet.range("l#{line}").Interior['ColorIndex'] =50
          $logger.log("Pass")
          $logger.log("Test Description:"+testCaseNumber+":" +testCaseId+":"+ description)
          else 
             worksheet.Range("l#{line}").Value="Fail"
            worksheet.range("l#{line}").Interior['ColorIndex'] =27
           $logger.log("Fail") 
            $logger.log("Test Description:"+testCaseNumber+":" +testCaseId+":"+ description)
            
          end
             #$ie.goto(test_site)
             line.succ!
               
             end
          $ie.frame( :name, 'header' ).link( :id, 'form1:Header:layoutPanel2:layoutPanel1:logout' ).click 
          $ie.close
          workbook.save
           workbook.close
           excel.Quit   
    end
 end
 
 
 
 
 
 
 
 
 
 