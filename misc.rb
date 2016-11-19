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
   def test_SC9_Login
       division_Text_Logger
       test_site = 'http://localhost:8080/alpha-swipecard-web'
        $ie.goto(test_site)
        excel = WIN32OLE::new("excel.Application")
        workbook = excel.Workbooks.Open("c:\\SClogin.xls") # directory Path where the test data is located
        worksheet = workbook.WorkSheets(9)
        worksheet.Select
        line = '2'

        while
            test_id= worksheet.Range("a#{line}")["Value"]  
            requestor_username= worksheet.Range("b#{line}")["Value"]  
            requestor_password= worksheet.Range("c#{line}")["Value"]  
            navPanel=worksheet.Range("d#{line}")["Value"]  
            testValidator1 = worksheet.Range("e#{line}")["Value"]
            description1 = worksheet.Range("f#{line}")["Value"]
            testCaseId1 = worksheet.Range("g#{line}")["Value"]
            testCaseNumber1 = worksheet.Range("h#{line}")["Value"]
            testValidator2 = worksheet.Range("i#{line}")["Value"]
            description2 = worksheet.Range("j#{line}")["Value"]
            testCaseId2 = worksheet.Range("k#{line}")["Value"]
            testCaseNumber2 = worksheet.Range("l#{line}")["Value"]
            testValidator3 = worksheet.Range("m#{line}")["Value"] 
            description3= worksheet.Range("n#{line}")["Value"]
            testCaseId3 = worksheet.Range("o#{line}")["Value"]
            testCaseNumber3 = worksheet.Range("p#{line}")["Value"]
            testValidator4 = worksheet.Range("q#{line}")["Value"] 
            description4= worksheet.Range("r#{line}")["Value"]
            testCaseId4 = worksheet.Range("s#{line}")["Value"]
            testCaseNumber4 = worksheet.Range("t#{line}")["Value"]
            testValidator5 = worksheet.Range("u#{line}")["Value"] 
            description5= worksheet.Range("v#{line}")["Value"]
            testCaseId5 = worksheet.Range("w#{line}")["Value"]
            testCaseNumber5 = worksheet.Range("x#{line}")["Value"]
            testValidator6 = worksheet.Range("y#{line}")["Value"] 
            description6= worksheet.Range("z#{line}")["Value"]
            testCaseId6 = worksheet.Range("aa#{line}")["Value"]
            testCaseNumber6 = worksheet.Range("ab#{line}")["Value"]
            testValidator7 = worksheet.Range("ac#{line}")["Value"] 
            testValidator8 = worksheet.Range("ad#{line}")["Value"] 
            description7= worksheet.Range("ae#{line}")["Value"]
            testCaseId7 = worksheet.Range("af#{line}")["Value"]
            testCaseNumber7 = worksheet.Range("ag#{line}")["Value"]
             description8= worksheet.Range("ah#{line}")["Value"]
            testCaseId8 = worksheet.Range("ai#{line}")["Value"]
            testCaseNumber8 = worksheet.Range("aj#{line}")["Value"]
            paginator1= worksheet.Range("ak#{line}")["Value"]
            paginator2= worksheet.Range("al#{line}")["Value"]
            paginator3= worksheet.Range("am#{line}")["Value"]
            paginator4= worksheet.Range("an#{line}")["Value"]
            paginator5= worksheet.Range("ao#{line}")["Value"]
            paginator6= worksheet.Range("ap#{line}")["Value"]
            paginator7= worksheet.Range("aq#{line}")["Value"]
            $logger.log(" ")
            $logger.log(test_id)
          
            $ie.text_field(:name,"j_username").set(requestor_username)
            $logger.log("Action: Entered " + requestor_username + " as username")
            $ie.text_field(:name,"j_password").set(requestor_password)
            $logger.log("Action: Entered " + requestor_password + " as password")
            $ie.button(:value, "Ok").click
            $logger.log( "Action: Clicked Login Button")
            
            $ie.frame( :name, 'header' ).link( :id, navPanel ).click
            
            $ie.frame( :name, 'main' ).image( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal:_datePicker_datePickerLink_image' ).click
            $ie.frame( :name, 'main' ).select_list( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal:_datePicker:monthMenu_list' ).select( 'February' )
            $ie.frame( :name, 'main' ).image( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal:_datePicker:previousMonthLink_image' ).click
            $ie.frame( :name, 'main' ).image( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal:_datePicker:previousMonthLink_image' ).click
            $ie.frame( :name, 'main' ).image( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal:_datePicker:previousMonthLink_image' ).click
            $ie.frame( :name, 'main' ).image( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal:_datePicker:previousMonthLink_image' ).click
            $ie.frame( :name, 'main' ).image( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal:_datePicker:nextMonthLink_image' ).click
            $ie.frame( :name, 'main' ).link( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal:_datePicker_link:6' ).click
            $ie.frame( :name, 'main' ).image( :id, 'form1:layoutPanel2:groupPanel1:dateToCal:_datePicker_datePickerLink_image' ).click
            $ie.frame( :name, 'main' ).select_list( :id, 'form1:layoutPanel2:groupPanel1:dateToCal:_datePicker:monthMenu_list' ).select( 'February' )
            $ie.frame( :name, 'main' ).link( :id, 'form1:layoutPanel2:groupPanel1:dateToCal:_datePicker_link:19' ).click
            $ie.frame( :name, 'main' ).button( :id, 'form1:layoutPanel2:searchBtn' ).click
           

            if($ie.frame( :name, 'main').contains_text( testValidator1 ))
              $logger.log("Pass")
              $logger.log("Test Description:"+testCaseNumber1+":" +testCaseId1+":"+ description1)
              
            else 
              $logger.log("Fail") 
              $logger.log("Test Description:"+testCaseNumber1+":" +testCaseId1+":"+ description1)
             
          end

#mark
            $ie.frame( :name, 'main' ).image( :id, paginator1).click
           # e.frame( :name, 'main' ).image( :id, 'form1:shiftInOutSummaryTbl:_tableActionsBottom:_paginationNextButton:_paginationNextButton_image' ).click
             if($ie.frame( :name, 'main').contains_text( testValidator2 ))
              $logger.log("Pass")
              $logger.log("Test Description:"+testCaseNumber2+":" +testCaseId2+":"+ description2)
              
            else 
              $logger.log("Fail") 
              $logger.log("Test Description:"+testCaseNumber2+":" +testCaseId2+":"+ description2)
              
            end
  #mark          
            $ie.frame( :name, 'main' ).image( :id, paginator2 ).click
            
              if($ie.frame( :name, 'main').contains_text( testValidator3 ))
              $logger.log("Pass")
              $logger.log("Test Description:"+testCaseNumber3+":" +testCaseId3+":"+ description3)
              
            else 
              $logger.log("Fail") 
              $logger.log("Test Description:"+testCaseNumber3+":" +testCaseId3+":"+ description3)
             
            end
            
          #  $ie.frame( :name, 'main' ).text_field( :id, 'form1:shiftInOutSummary:_tableActionsBottom:_paginationPageField_field' ).set( '27' )
            $ie.frame( :name, 'main' ).image( :id, paginator3 ).click
  
            if($ie.frame( :name, 'main').contains_text( testValidator4 ))
              $logger.log("Pass")
              $logger.log("Test Description:"+testCaseNumber4+":" +testCaseId4+":"+ description4)
              
            else 
              $logger.log("Fail") 
              $logger.log("Test Description:"+testCaseNumber4+":" +testCaseId4+":"+ description4)
             
            end
  
            $ie.frame( :name, 'main' ).image( :id, paginator4 ).click
            
              if($ie.frame( :name, 'main').contains_text( testValidator5))
              $logger.log("Pass")
              $logger.log("Test Description:"+testCaseNumber5+":" +testCaseId5+":"+ description5)
              
            else 
              $logger.log("Fail") 
              $logger.log("Test Description:"+testCaseNumber5+":" +testCaseId5+":"+ description5)
             
            end
            
            
            $ie.frame( :name, 'main' ).text_field( :id, paginator5 ).set( '3' )
            $ie.frame( :name, 'main' ).button( :id, paginator6 ).click

              if($ie.frame( :name, 'main').contains_text( testValidator6))
              $logger.log("Pass")
              $logger.log("Test Description:"+testCaseNumber6+":" +testCaseId6+":"+ description6)
             else 
              $logger.log("Fail") 
              $logger.log("Test Description:"+testCaseNumber6+":" +testCaseId6+":"+ description6)
              
            end

            $ie.frame( :name, 'main' ).image( :id, paginator7).click

             
            
          if($ie.frame( :name, 'main').contains_text( testValidator7 ) && $ie.frame( :name, 'main').contains_text( testValidator8 ))
            $logger.log("Pass")
            $logger.log("Test Description:"+testCaseNumber7+":" +testCaseId7+":"+ description7)
           
          else 
            $logger.log("Fail") 
            $logger.log("Test Description:"+testCaseNumber7+":" +testCaseId7+":"+ description7)
          end
          
          $ie.frame( :name, 'main' ).text_field( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal_field' ).set('02/20/2009')
         
          $ie.frame( :name, 'main' ).text_field( :id, 'form1:layoutPanel2:groupPanel1:dateToCal_field' ).set('02/20/2009')
          $ie.frame( :name, 'main' ).button( :value, 'Clear' ).click
             if($ie.frame( :name, 'main' ).text_field( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal_field' ).set(' ')&&$ie.frame( :name, 'main' ).text_field( :id, 'form1:layoutPanel2:groupPanel1:dateToCal_field' ).set(' '))
             
            $logger.log("Pass")
            $logger.log("Test Description:"+testCaseNumber8+":" +testCaseId8+":"+ description8)

           
          else 
            $logger.log("Fail") 
            $logger.log("Test Description:"+testCaseNumber8+":" +testCaseId8+":"+ description8)
          end
          $ie.frame( :name, 'header' ).link( :id, 'form1:Header:layoutPanel2:layoutPanel1:logout' ).click 
             #$ie.goto(test_site)
             line.succ!
               
             end
             
       ########################################################
# Test single search       
             line = '8'
             
             while
            test_id= worksheet.Range("a#{line}")["Value"]  
            requestor_username= worksheet.Range("b#{line}")["Value"]  
            requestor_password= worksheet.Range("c#{line}")["Value"]  
            navPanel=worksheet.Range("d#{line}")["Value"]  
            testValidator1 = worksheet.Range("e#{line}")["Value"]
            description1 = worksheet.Range("f#{line}")["Value"]
            testCaseId1 = worksheet.Range("g#{line}")["Value"]
            testCaseNumber1 = worksheet.Range("h#{line}")["Value"]
            testValidator2 = worksheet.Range("i#{line}")["Value"]
            description2 = worksheet.Range("j#{line}")["Value"]
            testCaseId2 = worksheet.Range("k#{line}")["Value"]
            testCaseNumber2 = worksheet.Range("l#{line}")["Value"]
            testValidator3 = worksheet.Range("m#{line}")["Value"] 
            description3= worksheet.Range("n#{line}")["Value"]
            testCaseId3 = worksheet.Range("o#{line}")["Value"]
            testCaseNumber3 = worksheet.Range("p#{line}")["Value"]
            testValidator4 = worksheet.Range("q#{line}")["Value"] 
            description4= worksheet.Range("r#{line}")["Value"]
            testCaseId4 = worksheet.Range("s#{line}")["Value"]
            testCaseNumber4 = worksheet.Range("t#{line}")["Value"]
            testValidator5 = worksheet.Range("u#{line}")["Value"] 
            description5= worksheet.Range("v#{line}")["Value"]
            testCaseId5 = worksheet.Range("w#{line}")["Value"]
            testCaseNumber5 = worksheet.Range("x#{line}")["Value"]
            testValidator6 = worksheet.Range("y#{line}")["Value"] 
            description6= worksheet.Range("z#{line}")["Value"]
            testCaseId6 = worksheet.Range("aa#{line}")["Value"]
            testCaseNumber6 = worksheet.Range("ab#{line}")["Value"]
            testValidator7 = worksheet.Range("ac#{line}")["Value"] 
            testValidator8 = worksheet.Range("ad#{line}")["Value"] 
            description7= worksheet.Range("ae#{line}")["Value"]
            testCaseId7 = worksheet.Range("af#{line}")["Value"]
            testCaseNumber7 = worksheet.Range("ag#{line}")["Value"]
             description8= worksheet.Range("ah#{line}")["Value"]
            testCaseId8 = worksheet.Range("ai#{line}")["Value"]
            testCaseNumber8 = worksheet.Range("aj#{line}")["Value"]
            paginator1= worksheet.Range("ak#{line}")["Value"]
            paginator2= worksheet.Range("al#{line}")["Value"]
            paginator3= worksheet.Range("am#{line}")["Value"]
            paginator4= worksheet.Range("an#{line}")["Value"]
            paginator5= worksheet.Range("ao#{line}")["Value"]
            paginator6= worksheet.Range("ap#{line}")["Value"]
            paginator7= worksheet.Range("aq#{line}")["Value"]
            $logger.log(" ")
            $logger.log(test_id)
          
            $ie.text_field(:name,"j_username").set(requestor_username)
            $logger.log("Action: Entered " + requestor_username + " as username")
            $ie.text_field(:name,"j_password").set(requestor_password)
            $logger.log("Action: Entered " + requestor_password + " as password")
            $ie.button(:value, "Ok").click
            $logger.log( "Action: Clicked Login Button")
            
            $ie.frame( :name, 'header' ).link( :id, navPanel ).click
            
           $ie.frame( :name, 'main' ).image( :id, 'form1:SingleUserSearchParameters:layoutPanel1:dateGroupPanel:dateFromCal:_datePicker_datePickerLink_image' ).click
          $ie.frame( :name, 'main' ).select_list( :id, 'form1:SingleUserSearchParameters:layoutPanel1:dateGroupPanel:dateFromCal:_datePicker:monthMenu_list' ).select( 'November' )
          $ie.frame( :name, 'main' ).select_list( :id, 'form1:SingleUserSearchParameters:layoutPanel1:dateGroupPanel:dateFromCal:_datePicker:yearMenu_list' ).select( '2008' )
          $ie.frame( :name, 'main' ).link( :id, 'form1:SingleUserSearchParameters:layoutPanel1:dateGroupPanel:dateFromCal:_datePicker_link:6' ).click
          $ie.frame( :name, 'main' ).image( :id, 'form1:SingleUserSearchParameters:layoutPanel1:dateGroupPanel:dateToCal:_datePicker_datePickerLink_image' ).click
          $ie.frame( :name, 'main' ).select_list( :id, 'form1:SingleUserSearchParameters:layoutPanel1:dateGroupPanel:dateToCal:_datePicker:monthMenu_list' ).select( 'February' )
          $ie.frame( :name, 'main' ).link( :id, 'form1:SingleUserSearchParameters:layoutPanel1:dateGroupPanel:dateToCal:_datePicker_link:19' ).click
        #  $ie.frame( :name, 'main' ).link( :id, 'form1:SingleUserSearchParameters:layoutPanel2:groupPanel1:dateToCal:_datePicker_link:19' ).click  
           $ie.frame( :name, 'main' ).button( :id, 'form1:SingleUserSearchParameters:layoutPanel1:searchButton' ).click
         #  ie.frame( :name, 'main' ).image( :id, 'form1:SingleUserSearchParameters:layoutPanel1:dateGroupPanel:dateToCal:_datePicker_datePickerLink_image' ).click
#ie.frame( :name, 'main' ).select_list( :id, 'form1:SingleUserSearchParameters:layoutPanel1:dateGroupPanel:dateToCal:_datePicker:monthMenu_list' ).select( 'February' )

            if($ie.frame( :name, 'main').contains_text( testValidator1 ))
              $logger.log("Pass")
              $logger.log("Test Description:"+testCaseNumber1+":" +testCaseId1+":"+ description1)
              
            else 
              $logger.log("Fail") 
              $logger.log("Test Description:"+testCaseNumber1+":" +testCaseId1+":"+ description1)
             
          end

#mark
            $ie.frame( :name, 'main' ).image( :id, paginator1).click
           # e.frame( :name, 'main' ).image( :id, 'form1:shiftInOutSummaryTbl:_tableActionsBottom:_paginationNextButton:_paginationNextButton_image' ).click
             if($ie.frame( :name, 'main').contains_text( testValidator2 ))
              $logger.log("Pass")
              $logger.log("Test Description:"+testCaseNumber2+":" +testCaseId2+":"+ description2)
              
            else 
              $logger.log("Fail") 
              $logger.log("Test Description:"+testCaseNumber2+":" +testCaseId2+":"+ description2)
              
            end
  #mark          
            $ie.frame( :name, 'main' ).image( :id, paginator2 ).click
            
              if($ie.frame( :name, 'main').contains_text( testValidator3 ))
              $logger.log("Pass")
              $logger.log("Test Description:"+testCaseNumber3+":" +testCaseId3+":"+ description3)
              
            else 
              $logger.log("Fail") 
              $logger.log("Test Description:"+testCaseNumber3+":" +testCaseId3+":"+ description3)
             
            end
            
          #  $ie.frame( :name, 'main' ).text_field( :id, 'form1:shiftInOutSummary:_tableActionsBottom:_paginationPageField_field' ).set( '27' )
            $ie.frame( :name, 'main' ).image( :id, paginator3 ).click
  
            if($ie.frame( :name, 'main').contains_text( testValidator4 ))
              $logger.log("Pass")
              $logger.log("Test Description:"+testCaseNumber4+":" +testCaseId4+":"+ description4)
              
            else 
              $logger.log("Fail") 
              $logger.log("Test Description:"+testCaseNumber4+":" +testCaseId4+":"+ description4)
             
            end
  
            $ie.frame( :name, 'main' ).image( :id, paginator4 ).click
            
              if($ie.frame( :name, 'main').contains_text( testValidator5))
              $logger.log("Pass")
              $logger.log("Test Description:"+testCaseNumber5+":" +testCaseId5+":"+ description5)
              
            else 
              $logger.log("Fail") 
              $logger.log("Test Description:"+testCaseNumber5+":" +testCaseId5+":"+ description5)
             
            end
            
            
            $ie.frame( :name, 'main' ).text_field( :id, paginator5 ).set( '3' )
            $ie.frame( :name, 'main' ).button( :id, paginator6 ).click

              if($ie.frame( :name, 'main').contains_text( testValidator6))
              $logger.log("Pass")
              $logger.log("Test Description:"+testCaseNumber6+":" +testCaseId6+":"+ description6)
             else 
              $logger.log("Fail") 
              $logger.log("Test Description:"+testCaseNumber6+":" +testCaseId6+":"+ description6)
              
            end

            $ie.frame( :name, 'main' ).image( :id, paginator7).click

             
            
          if($ie.frame( :name, 'main').contains_text( testValidator7 ) && $ie.frame( :name, 'main').contains_text( testValidator8 ))
            $logger.log("Pass")
            $logger.log("Test Description:"+testCaseNumber7+":" +testCaseId7+":"+ description7)
           
          else 
            $logger.log("Fail") 
            $logger.log("Test Description:"+testCaseNumber7+":" +testCaseId7+":"+ description7)
          end
          
          $ie.frame( :name, 'main' ).text_field( :id, 'form1:SingleUserSearchParameters:layoutPanel1:dateGroupPanel:dateFromCal_field' ).set('02/20/2009')
         
          $ie.frame( :name, 'main' ).text_field( :id, 'form1:SingleUserSearchParameters:layoutPanel1:dateGroupPanel:dateToCal_field' ).set('02/20/2009')
          $ie.frame( :name, 'main' ).button( :value, 'Clear' ).click
             if($ie.frame( :name, 'main' ).text_field( :id, 'form1:SingleUserSearchParameters:layoutPanel1:dateGroupPanel:dateFromCal_field' ).set(' ')&&$ie.frame( :name, 'main' ).text_field( :id, 'form1:SingleUserSearchParameters:layoutPanel1:dateGroupPanel:dateToCal_field').set(' '))
             
            $logger.log("Pass")
            $logger.log("Test Description:"+testCaseNumber8+":" +testCaseId8+":"+ description8)

           
          else 
            $logger.log("Fail") 
            $logger.log("Test Description:"+testCaseNumber8+":" +testCaseId8+":"+ description8)
          end
          $ie.frame( :name, 'header' ).link( :id, 'form1:Header:layoutPanel2:layoutPanel1:logout' ).click 
             #$ie.goto(test_site)
             line.succ!
               
             end
          #  $ie.frame( :name, 'header' ).link( :id, 'form1:Header:layoutPanel2:layoutPanel1:logout' ).click 
          $ie.close
          workbook.save
           workbook.close
           excel.Quit   
    end
 end