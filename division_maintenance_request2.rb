require 'watir'   # the controller
require 'win32ole'
require 'watir'
include Watir
#test::unit includes
require 'watir/clickJSDialog.rb'
require 'test/unit' 
require 'test/unit/ui/console/testrunner'

class TC_MAINTENANCE_REQUESTs < Test::Unit::TestCase
       def test_division_maintenance
       # division_maintenance_start
       $ie = IE.new
        test_site = 'http://gondor:8080/pim2/login.action'
        $ie.goto(test_site)
        excel = WIN32OLE::new("excel.Application")
        workbook = excel.Workbooks.Open("c:\\division maintenance request.xls") # directory Path where the test data is located
        worksheet = workbook.WorkSheets(1)
        worksheet.Select
        line = '2'

        while
            test_id= worksheet.Range("a#{line}")["Value"]  
            requestor_username= worksheet.Range("b#{line}")["Value"]  
            requestor_password= worksheet.Range("c#{line}")["Value"]  
            search_criteria = worksheet.Range("d#{line}")["Value"]
            item_id = worksheet.Range("e#{line}")["Value"]
            manufacturer_lookup = worksheet.Range("f#{line}")["Value"]
            manufacturer_lookup_index = worksheet.Range("g#{line}")["Value"]
            brand_lookup = worksheet.Range("h#{line}")["Value"]
            brand_lookup_index = worksheet.Range("i#{line}")["Value"]
            approver_username= worksheet.Range("j#{line}")["Value"]  
            approver_password= worksheet.Range("k#{line}")["Value"]  
           # $logger.log(" ")
            #$logger.log(test_id)
            $ie.text_field(:name,"userForm.userName").set(requestor_username)
            #$logger.log("Action: Entered " + requestor_username + " as username")
            $ie.text_field(:name,"userForm.userPassword").set(requestor_password)
            #$logger.log("Action: Entered " + requestor_password + " as password")
            $ie.button(:value, "Login").click
            #$logger.log( "Action: Clicked Login Button")
            $ie.image(:index, 2).click
            $ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click
          #  if (worksheet.Range("d#{line}")["Value"]  ==nil)
            #    worksheet.Range("d#{line}")["Value"]=nil
            #else
              #  $ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.searchKeywords").set(search_criteria)
              #  $logger.log("Action: Entered " +search_criteria + " in the Product Description Input Field")
           # end
            #$ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click
            #$logger.log( "Action: Clicked Search Button Button")
            #if($ie.pageContainsText("Found PIM Groups"))
              #  $ie.button(:name, "method:executeShowAll").click
              #  $logger.log("Action: Clicked Show All Button from the Search Summary Page")
               # sleep(5)
                #$ie.checkbox(:id, "#{item_id}").set
            #else
              #  $ie.checkbox(:id, "#{item_id}").set
            #end
            #$ie.button(:name, "itemMaintenance").click
            #$logger.log("Action: Clicked Item Maintenance Button")
            #sleep(5)
            #$ie2= Watir::IE.attach(:url, 'http://gondor:8080/pim2/maintenanceRequest.action')
            #$ie2.text_field(:name,"itemRequestVo.descriptionNote").set("diomar")
           # $ie2.text_field(:name,"itemVo.mfrName").fire_event("onDblClick")
            #sleep(2)
           # $logger.log("Action: Double Click The Mfr Look Up Text Field")
            #$ie3= Watir::IE.attach(:url, 'http://gondor:8080/pim2/manufacturerLookup.action')
            #$ie3.text_field(:name,"usfManufacturerName").set(manufacturer_lookup)
            #$logger.log("Action: Entered " + manufacturer_lookup+ " in the Mfr Name Input Field")
            #$ie3.button(:name, "method:search").click
            #sleep(2)
            #$ie3.link(:index, "#{manufacturer_lookup_index}").click
            #$ie2.text_field(:name,"itemVo.brandName").fire_event("onDblClick")
            #sleep(2)
           # $ie4= Watir::IE.attach(:url, 'http://gondor:8080/pim2/brandLookup.action')
            #$ie4.text_field(:id,"brandName").set(brand_lookup)
            #$logger.log("Action: Entered " + brand_lookup+ " in the Brand Input Field")
            #$ie4.button(:name, "method:search").click
            #sleep(2)
           # $ie4.link(:index, "#{brand_lookup_index}").click
            #$ie2.button(:name, "method:save").click
            #sleep(2)
            #$ie2.button(:value, "Close").click
            $ie.image(:index, 1).click
            $ie.link(:index, "47").click
            sleep(2)
            $ie.link(:index, "49").click
            $ie.checkboxes[1].set
            $ie.button(:value, "Open").click
 
          
          gago=getWindowTitle(-1)
          puts (gago)
           # test_site2 = 'http://gondor:8080/pim2/login.action'
          #$ie.goto(test_site2)
            sleep(2)
            #title=$ie.getWindowTitle(hWnd) 
            #$logger.log(title)
            #$logger.log($ie.getWindowTitle)
            #$logger.log($ie.title($ie.front?))

           # $logger.log(url)
            
            #$logger.log(" ")
            
            line.succ!
        end
        $ie.close
    end
  end
          
 
         