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
   def test_SC1_Login
       division_Text_Logger
       test_site = 'http://localhost:8080/alpha-swipecard-web'
        $ie.goto(test_site)
        excel = WIN32OLE::new("excel.Application")
        workbook = excel.Workbooks.Open("c:\\SClogin.xls") # directory Path where the test data is located
        worksheet = workbook.WorkSheets(8)
        worksheet.Select
        line = '2'

      $ie.text_field( :name, 'j_username' ).set( 'crcortez' )
       $ie.text_field( :name, 'j_password' ).set( 'onesecret' )
        $ie.button( :name, 'login' ).click

        $ie.frame( :name, 'header' ).link( :id, 'form1:Navigation:navPnl:shift' ).click

 while
            test_id= worksheet.Range("a#{line}")["Value"]  
            dateFromRange= worksheet.Range("b#{line}")["Value"]  
            dateToRange= worksheet.Range("c#{line}")["Value"]  
            testValidator1 = worksheet.Range("d#{line}")["Value"]
            testValidator2 = worksheet.Range("e#{line}")["Value"]
            description = worksheet.Range("f#{line}")["Value"]
            testCaseId = worksheet.Range("g#{line}")["Value"]
            testCaseNumber= worksheet.Range("h#{line}")["Value"]
            
            $logger.log(" ")
            $logger.log(test_id)
 
          $ie.frame( :name, 'main' ).text_field( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal_field' ).set(dateFromRange)
          $logger.log("Action: Entered " + dateFromRange + " as Date From")
          $ie.frame( :name, 'main' ).text_field( :id, 'form1:layoutPanel2:groupPanel1:dateToCal_field' ).set(dateToRange)
          $logger.log("Action: Entered " + dateToRange + " as Date To")
          $ie.frame( :name, 'main' ).button( :id, 'form1:layoutPanel2:searchBtn' ).click
          $ie.frame( :name, 'main' ).button( :id, 'form1:layoutPanel2:searchBtn' ).click
         
          if($ie.frame( :name, 'main').contains_text( testValidator1 ) && $ie.frame( :name, 'main').contains_text( testValidator2))
           worksheet.Range("i#{line}").Value="Pass"
            worksheet.range("i#{line}").Interior['ColorIndex'] =50
          $logger.log("Pass")
            $logger.log("Test Description:"+testCaseNumber+":" +testCaseId+":"+ description)
          else 
            worksheet.Range("i#{line}").Value="Fail"
            worksheet.range("i#{line}").Interior['ColorIndex'] =28
            $logger.log("Fail") 
            $logger.log("Test Description:"+testCaseNumber+":" +testCaseId+":"+ description)
            
          end
            line.succ!
               
             end
          
            ######################################################################           
  #Timesheet Correlation Codes
        line = '11'   
             
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
            worksheet.range("j#{line}").Interior['ColorIndex'] =28
            $logger.log("Fail") 
            $logger.log("Test Description:"+testCaseNumber+":" +testCaseId+":"+ description)
            
          end
             #$ie2.goto(test_site)
             line.succ!
               
             end
     ######################################################################           
  #Timesheet Annotation
        line = '25'   
    
       $ie.frame( :name, 'header' ).link( :id, 'form1:Navigation:navPnl:timesheet' ).click
        $logger.log("Timesheet Correlation") 
 while
            test_id= worksheet.Range("a#{line}")["Value"]  
            textEntry= worksheet.Range("b#{line}")["Value"]  
            testValidator1 = worksheet.Range("c#{line}")["Value"]
            description = worksheet.Range("d#{line}")["Value"]
            testCaseId = worksheet.Range("e#{line}")["Value"]
            testCaseNumber= worksheet.Range("f#{line}")["Value"]
            
            
            $logger.log(" ")
            $logger.log(test_id)
          $ie.frame( :name, 'header' ).link( :id, 'form1:Navigation:navPnl:timesheet' ).click
          $ie.frame( :name, 'main' ).text_field( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal_field' ).set('02/20/2009')
         
          $ie.frame( :name, 'main' ).text_field( :id, 'form1:layoutPanel2:groupPanel1:dateToCal_field' ).set('02/20/2009')
    
          $ie.frame( :name, 'main' ).text_field( :id, 'form1:layoutPanel2:groupPanel2:nameTxt_field' ).set(' ')
           
          $ie.frame( :name, 'main' ).button( :id, 'form1:layoutPanel2:searchBtn' ).click
 
          $ie.frame( :name, 'main' ).link( :id, 'form1:timesheetTable:tableRowGroup1:0:tableColumn6:swipeDetail' ).click
          $ie.frame( :name, 'main' ).link( :id, 'form1:table1:tableRowGroup1:0:tableColumn5:annotateLink' ).click
          
          $ie.frame(:name, 'main').text_field(:id, 'form1:taAnnotation_field').set(textEntry)
          $logger.log("Action: Entered " + textEntry + " as Annotation ")
  
          $ie.frame( :name, 'main' ).button( :id, 'form1:button1' ).click
          
          #ie.frame( :name, 'main' ).span( :id, 'form1:table1:tableRowGroup1:0:tableColumn4:staticText4' ).click
  if($ie.frame( :name, 'main').contains_text( testValidator1 ))
           worksheet.Range("g#{line}").Value="Pass"
            worksheet.range("g#{line}").Interior['ColorIndex'] =50
          $logger.log("Pass")
          $logger.log(testCaseNumber+":" +testCaseId+":"+ description)
          else 
            worksheet.Range("g#{line}").Value="Fail"
            worksheet.range("g#{line}").Interior['ColorIndex'] =28
           $logger.log("Fail") 
           $logger.log(testCaseNumber+":" +testCaseId+":"+ description)
            
          end
             #$ie2.goto(test_site)
             line.succ!
               
             end  
          $ie.frame( :name, 'header' ).link( :id, 'form1:Header:layoutPanel2:layoutPanel1:logout' ).click 
          $ie.close
           workbook.save
           #workbook.close
          excel.Quit   
    end
 end
 