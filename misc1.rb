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
#$ie = IE.new

#$ie.goto( 'http://localhost:8080/alpha-swipecard-web/faces/pages/index.jsp' )

$ie.text_field( :name, 'j_username' ).set( 'crcortez' )
$ie.text_field( :name, 'j_password' ).set( 'onesecret' )
$ie.button( :name, 'login' ).click

$ie.frame( :name, 'header' ).link( :id, 'form1:Navigation:navPnl:audit' ).click

$ie.frame( :name, 'main' ).image( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal:_datePicker_datePickerLink_image' ).click
$ie.frame( :name, 'main' ).select_list( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal:_datePicker:monthMenu_list' ).select( 'February' )
$ie.frame( :name, 'main' ).link( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal:_datePicker_link:25' ).click
$ie.frame( :name, 'main' ).image( :id, 'form1:layoutPanel2:groupPanel1:dateToCal:_datePicker_datePickerLink_image' ).click
$ie.frame( :name, 'main' ).select_list( :id, 'form1:layoutPanel2:groupPanel1:dateToCal:_datePicker:monthMenu_list' ).select( 'February' )
$ie.frame( :name, 'main' ).link( :id, 'form1:layoutPanel2:groupPanel1:dateToCal:_datePicker_link:25' ).click
$ie.frame( :name, 'main' ).button( :id, 'form1:layoutPanel2:searchBtn' ).click
  
  if($ie.frame( :name, 'main').contains_text( "ShiftSingleUserSearch" ))
              $logger.log("Pass")
              $logger.log("Test Description:TC1:SC_Search_Audit_Trail_Valid_Date_Fields_Calendar_05_Pos:System will be able to search based on date fields and and with frequency parameter.")
              
            else 
              $logger.log("Fail") 
               $logger.log("Test Description:TC1:SC_Search_Audit_Trail_Valid_Date_Fields_Calendar_05_Pos:System will be able to search based on date fields and and with frequency parameter.")
              
          end
$ie.frame( :name, 'main' ).button( :value, 'Clear' ).click

 if($ie.frame( :name, 'main' ).text_field( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal_field' ).set(' ')&&$ie.frame( :name, 'main' ).text_field( :id, 'form1:layoutPanel2:groupPanel1:dateToCal_field' ).set(' '))
             
            $logger.log("Pass")
            $logger.log("Test Description:TC4:SC_Search_Audit_Trail_Valid_All_Fields_Reset_01_Pos:All text fields should cleared out.")

           
          else 
            $logger.log("Fail") 
            $logger.log("Test Description:TC4:SC_Search_Audit_Trail_Valid_All_Fields_Reset_01_Pos:All text fields should cleared out.")
          end
          
$ie.frame( :name, 'header' ).link( :id, 'form1:Header:layoutPanel2:layoutPanel1:logout' ).click


$ie.text_field( :name, 'j_username' ).set( 'cgmanuel' )
$ie.text_field( :name, 'j_password' ).set( 'kilouwa31' )
$ie.button( :name, 'login' ).click

$ie.frame( :name, 'header' ).link( :id, 'form1:Navigation:navPnl:whoIsHere' ).click

$ie.frame( :name, 'main' ).image( :id, 'whosHereForm:datePicker:_datePicker_datePickerLink_image' ).click
$ie.frame( :name, 'main' ).select_list( :id, 'whosHereForm:datePicker:_datePicker:monthMenu_list' ).select( 'February' )
$ie.frame( :name, 'main' ).link( :id, 'whosHereForm:datePicker:_datePicker_link:25' ).click
$ie.frame( :name, 'main' ).button( :id, 'whosHereForm:searchBtn' ).click
$ie.frame( :name, 'main' ).select_list( :id, 'whosHereForm:reportType_list' ).select( "Who's Out" )
$ie.frame( :name, 'main' ).button( :id, 'whosHereForm:searchBtn' ).click

if($ie.frame( :name, 'main').contains_text( "Aldrich Veluz" ))
              $logger.log("Pass")
              $logger.log("Test Description:TC1:SC_Search_Who's_Here_Valid_Fields_Calendar_Who_Is_Out_04_Pos:System will be able to search with all valid field.Who is out criteria.Calendar icon")
              
            else 
              $logger.log("Fail") 
               $logger.log("Test Description:TC1:SC_Search_Who's_Here_Valid_Fields_Calendar_Who_Is_Out_04_Pos:System will be able to search with all valid field.Who is out criteria.Calendar icon")
          end


$ie.frame( :name, 'main' ).image( :id, 'whosHereForm:dataGrid:_tableActionsBottom:_paginationNextButton:_paginationNextButton_image' ).click
 if($ie.frame( :name, 'main').contains_text("Michelle delosSantos"))
              $logger.log("Pass")
              $logger.log("Test Description:TC6:SC_Search_Who's_Here_Results_Table_Next_Page_01_Pos:Verify Next Page")
              
            else 
              $logger.log("Fail") 
              $logger.log("Test Description:TC6:SC_Search_Who's_Here_Results_Table_Next_Page_01_Pos:Verify Next Page")
              
            end

#$ie.frame( :name, 'main' ).text_field( :id, 'whosHereForm:dataGrid:_tableActionsBottom:_paginationPageField_field' ).set( '2' )

$ie.frame( :name, 'main' ).image( :id, 'whosHereForm:dataGrid:_tableActionsBottom:_paginationPrevButton:_paginationPrevButton_image' ).click
 if($ie.frame( :name, 'main').contains_text(  "Aldrich Veluz" ))
              $logger.log("Pass")
              $logger.log("Test Description:TC6:SC_Search_Who's_Here_Results_Table_Prev_Page_03_Pos:Verify Previous Page")
              
            else 
              $logger.log("Fail") 
              $logger.log("Test Description:TC6:SC_Search_Who's_Here_Results_Table_Prev_Page_03_Pos:Verify Previous Page")
              
            end

#$ie.frame( :name, 'main' ).text_field( :id, 'whosHereForm:dataGrid:_tableActionsBottom:_paginationPageField_field' ).set( '1' )
$ie.frame( :name, 'main' ).image( :id, 'whosHereForm:dataGrid:_tableActionsBottom:_paginationLastButton:_paginationLastButton_image' ).click
 if($ie.frame( :name, 'main').contains_text( "Michelle delosSantos" ))
              $logger.log("Pass")
              $logger.log("Test Description:TC6:SC_Search_Who's_Here_Results_Table_Last_Page_02_Pos:Verify Last Page")
              
            else 
              $logger.log("Fail") 
               $logger.log("Test Description:TC6:SC_Search_Who's_Here_Results_Table_Last_Page_02_Pos:Verify Last Page")
              
            end
#$ie.frame( :name, 'main' ).text_field( :id, 'whosHereForm:dataGrid:_tableActionsBottom:_paginationPageField_field' ).set( '2' )
$ie.frame( :name, 'main' ).image( :id, 'whosHereForm:dataGrid:_tableActionsBottom:_paginationFirstButton:_paginationFirstButton_image' ).click
 if($ie.frame( :name, 'main').contains_text( "Aldrich Veluz" ))
              $logger.log("Pass")
              $logger.log("Test Description:TC6:SC_Search_Who's_Here_Results_Table_First_Page_04_Pos:Verify First Page")
              
            else 
              $logger.log("Fail") 
                 $logger.log("Test Description:TC6:SC_Search_Who's_Here_Results_Table_First_Page_04_Pos:Verify First Page")
             
            end
$ie.frame( :name, 'main' ).text_field( :id, 'whosHereForm:dataGrid:_tableActionsBottom:_paginationPageField_field' ).set( '2' )
$ie.frame( :name, 'main' ).button( :id, 'whosHereForm:dataGrid:_tableActionsBottom:_paginationSubmitButton' ).click
 if($ie.frame( :name, 'main').contains_text( "Michelle delosSantos"))
              $logger.log("Pass")
              $logger.log("Test Description:TC6:SC_Search_Single_Who's_Here_Results_Table_Page_05_Pos:Verify that items in specified page in page search is correct")
              
            else 
              $logger.log("Fail") 
              $logger.log("Test Description:TC6:SC_Search_Single_Who's_Here_Results_Table_Page_05_Pos:Verify that items in specified page in page search is correct")
              
            end

$ie.frame( :name, 'main' ).image( :id, 'whosHereForm:dataGrid:_tableActionsBottom:_paginateButton:_paginateButton_image' ).click
if($ie.frame( :name, 'main').contains_text( "Aldrich Veluz" )&&$ie.frame( :name, 'main').contains_text( "Michelle delosSantos"))
              $logger.log("Pass")
               $logger.log("Test Description:TC6:SC_Search_Single_Who's_Here_Results_Table_View_All_06_Pos:View all data in one page")
                            
            else 
              $logger.log("Fail") 
              $logger.log("Test Description:TC6:SC_Search_Single_Who's_Here_Results_Table_View_All_06_Pos:View all data in one page")
              
            end
$ie.frame( :name, 'main' ).button( :value, 'Clear' ).click
 if($ie.frame( :name, 'main' ).text_field( :id, 'whosHereForm:datePicker_field' ).set(' ')&&$ie.frame( :name, 'main' ).text_field( :id, 'whosHereForm:hour_field' ).set(' ')&&$ie.frame( :name, 'main' ).text_field( :id, 'whosHereForm:min_field' ).set(' '))
             
            $logger.log("Pass")
            $logger.log("Test Description:TC5:SC_Search_Who's_Here_Valid_All_Fields_Reset_01_Pos:All text fields should cleared out.")

           
          else 
            $logger.log("Fail") 
            $logger.log("Test Description:TC5:SC_Search_Who's_Here_Valid_All_Fields_Reset_01_Pos:All text fields should cleared out.")
          end
$ie.frame( :name, 'header' ).link( :id, 'form1:Header:layoutPanel2:layoutPanel1:logout' ).click
workbook.save
           workbook.close
           excel.Quit   
  end
 end