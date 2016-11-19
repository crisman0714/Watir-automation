require 'watir'   # the controller
require 'win32ole'
include Watir

#test::unit includes
require 'test/unit' 
require 'test/unit/ui/console/testrunner'
mesg ="test"
text = "test text"

test_site1 = 'http://gondor:8080/pim2/login.action'
ie = Watir::IE.new
excel = WIN32OLE::new("excel.Application")
workbook = excel.Workbooks.Open("c:\\1search.xls")
worksheet = workbook.WorkSheets(1) # get first workbook
worksheet.Select

puts "PIM2 Testing"
puts "  "
ie.goto(test_site1)
ie.text_field(:name, "userForm.userName").set(worksheet.Range("b1").value)       # q is the name of the search field
ie.text_field(:name, "userForm.userPassword").set(worksheet.Range("b2").value)       # q is the name of the search field
ie.button(:value, "Login").click 
##Click Search Button
puts "Click Search Button."
ie.image(:index, 2).click

##Search Using Product Description As the Search Criteria
puts "  "
puts "Search by description." 
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.searchKeywords").set(worksheet.Range("b3").value)
ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click
if (ie.pageContainsText("0 items found."))
    count=ie.checkboxes.length
    puts("Product Description: " +worksheet.Range("b3").value + " #{count}" + " item(s) found")
    puts("Test: Search by Description - FAILED")
    ie.button(:name, "method:cancel").click
else
    if(ie.pageContainsText("Found PIM Groups"))
        ie.button(:name, "method:executeShowAll").click
        sleep(2)
        count=ie.checkboxes.length
        puts("Product Description: " +worksheet.Range("b3").value + " #{count}" + " item(s) found")
        puts("Test: Search by Description - PASSED")
        ie.image(:index, 2).click
    else
        sleep(2)
        count=ie.checkboxes.length
        puts("Product Description: " +worksheet.Range("b3").value + " #{count}" + " item(s) found")
        puts("Test: Search by Description - PASSED")
        ie.button(:name, "method:cancel").click
    end
end  

##Go Back To Home Page
puts "  "
puts "Go To Home Page."
ie.image(:index, 1).click

## Test Log Out Buttton
puts "  "
puts "Log Out To PIM."
ie.image(:index, 3).click
  
### Re Log IN
puts "  "
puts "Re LOgin To PIM"
ie.text_field(:name, "userForm.userName").set(worksheet.Range("b4").value)     
ie.text_field(:name, "userForm.userPassword").set(worksheet.Range("b5").value)      
ie.button(:value, "Login").click   # "btnG" is the name of the Search button
ie.image(:index, 2).click
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click
##############################################
##############Test Count Button #####################
##############################################
puts "  "
puts "Test Count Button" 
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.searchKeywords").set(worksheet.Range("b6").value)
ie.frame("main").frame(:name,"query").button(:value, "count").click
if (ie.pageContainsText("0 Products found "))
    puts("Count Button: " +worksheet.Range("b6").value + " 0 item(s) found")
    puts("TEST: Count Button - FAILED")
else
    puts("Count: " +worksheet.Range("b6").value + " found")
    puts("TEST: Count Button - PASSED")
end
sleep(2)

###############################################
#############Test Cancel Button - Search Count Result Page###########
################################################
puts "  "
puts "Click Cancel Button - Search Count Result Page"
ie.button(:name, 'Cancel').click
sleep(2)
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click
##############################################
##############test Reset Functionality##########################
###############################################
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.pimGroupNumber").set(worksheet.Range("b10").value)
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.productNumberType").select("Psys#")
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.productNumber").set(worksheet.Range("b13").value)
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.brandName").set(worksheet.Range("b18").value)
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.usfManufacturerNumber").set(worksheet.Range("b20").value)
puts "  "
puts "test Click Reset Button - searchHome page"
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click

#############################################
###########Test Reset Button search Home Page################
#############################################
#ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click

################################################
###########Test Search using PIM Class as Search Criteria##########
################################################
puts "  "
puts "Search Criteria: PIM Class"
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.pimClassName").set(worksheet.Range("b8").value)
ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click

if (ie.pageContainsText("0 items found."))
    puts("PIM Class ID: " +worksheet.Range("b8").value + " 0 item(s) found")
    puts("Test: Search By PIM Class ID - FAILED")
    ie.button(:name, "method:cancel").click
else
    if(ie.pageContainsText("Found PIM Groups"))
        ie.button(:name, "method:executeShowAll").click
        sleep(2)
        count=ie.checkboxes.length
        puts("PIM Class ID: " +worksheet.Range("b8").value +  " #{count}" + " item(s) found")
        puts("Test: Searcg By PIM Class ID - PASSED")
        ie.image(:index, 2).click
    else
        count=ie.checkboxes.length
        puts("PIM Class ID: " +worksheet.Range("b8").value)
        puts("Test: Search By PIM Class ID - PASSED")
        ie.button(:name, "method:cancel").click
    end
end 

ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click
  
################################################
###########Test Search using PIM Category as Search Criteria##########
################################################
puts "  "
puts "Search Criteria: PIM Category"
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.pimCategory").set(worksheet.Range("b9").value)
ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click
if (ie.pageContainsText("0 items found."))
    puts("PIM Category ID: " +worksheet.Range("b9").value + " 0 item(s) found")
    puts("Test: PIM Category ID - FAILED")
    ie.button(:name, "method:cancel").click
else
    if(ie.pageContainsText("Found PIM Groups"))
        ie.button(:name, "method:executeShowAll").click
        sleep(2)
        count=ie.checkboxes.length
        puts("PIM Category ID: " +worksheet.Range("b9").value + " 0 item(s) found")
        puts("Test: PIM Category ID - PASSED")
        ie.image(:index, 2).click
    else
        sleep(2)
        count=ie.checkboxes.length
        puts("PIM Category ID: " +worksheet.Range("b9").value + " 0 item(s) found")
        puts("Test: PIM Category ID - PASSED")
        ie.button(:name, "method:cancel").click
    end
  end 
ie.image(:index, 2).click
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click

################################################
###########Test Search using PIM Group as Search Criteria##########
################################################
puts "Hello World"
puts "Search Criteria: PIM Group"
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.pimGroupNumber").set(worksheet.Range("b10").value)
ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click
puts "!!!!!!!!!!!!!!!!!MABUHAY!!!!"
if (ie.pageContainsText("0 items found."))
    puts("PIM Group ID: " +worksheet.Range("b10").value)
    puts("Test: PIM Group ID - FAILED")
    ie.button(:name, "method:cancel").click
else
    puts "pim grooup else"
    if(ie.pageContainsText("Found PIM Groups"))
        puts "pim group else if"
        ie.button(:name, "method:executeShowAll").click
        sleep(2)
        puts "diomar"
        count=ie.checkboxes.length
        puts("PIM Group ID: " +worksheet.Range("b10").value + " #{count}" + " item(s) found")
        puts("Test: Search by PIM Group ID - PASSED")
        ie.image(:index, 2).click
    else
        puts "pim group else else"
        sleep(2)
        puts "diomar1"
        count=ie.checkboxes.length
        puts("PIM Group ID: " +worksheet.Range("b10").value + " #{count}" + " item(s) found")
        puts("Test: Search by PIM Group ID - PASSED")
        ie.button(:name, "method:cancel").click
    end
end 
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click
###############################################
####Test Product Number Type Dropdown#############
##############################################
puts "  "
puts "Test Product Number Type Dropdown"
sleep(1)
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.productNumberType").select("Psys#")
sleep(1)
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.productNumberType").select("Asys#")
sleep(1)
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.productNumberType").select("UPC")
sleep(1)
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.productNumberType").select("USF Prod#")
sleep(1)
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.productNumberType").select("Prod Master#")
sleep(1)
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.productNumberType").select("GTIN")
sleep(1)
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.productNumberType").select("Mfr Prod#")
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click

####################################
########Search Using Mfr Prod # as search criteria####
#####################################
puts "  "
puts "Search Using Mfr Prod # as Search Criteria"
sleep(2)
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.productNumberType").select("Mfr Prod#")
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.productNumber").set(worksheet.Range("b11").value)
ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click
if (ie.pageContainsText("0 items found."))
    count=ie.checkboxes.length
    puts("USF Mfr Prod #: " +worksheet.Range("b11").value + " #{count}" + " item(s) found")
    puts("Test: USF Mfr Prod # - FAILED")
    ie.button(:name, "method:cancel").click
else
    if(ie.pageContainsText("Found PIM Groups"))
        ie.button(:name, "method:executeShowAll").click
        sleep(2)
        count=ie.checkboxes.length
        puts("USF Mfr Prod #: " +worksheet.Range("b11").value + " #{count}" + " item(s) found")
        puts("Test: USF Mfr Prod # - PASSED")
        ie.image(:index, 2).click
    else
        sleep(2)
        count=ie.checkboxes.length
        puts("USF Mfr Prod #: " +worksheet.Range("b11").value + " #{count}" + " item(s) found")
        puts("Test: USF Mfr Prod # - PASSED")
        ie.button(:name, "method:cancel").click
    end
end 
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click

####################################
########Search Using Psys # as search criteria####
#####################################
puts "  "
puts "Psys # as Search Criteria"
sleep(2)
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.productNumberType").select("Psys#")
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.productNumber").set(worksheet.Range("b12").value)
ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click
if (ie.pageContainsText("0 items found."))
    count=ie.checkboxes.length
    puts("PSYS #: " +worksheet.Range("b12").value+ " #{count}" + " item(s) found")
    puts("Test: PSYS # - FAILED")
    ie.button(:name, "method:cancel").click
else
    if(ie.pageContainsText("Found PIM Groups"))
        ie.button(:name, "method:executeShowAll").click
        sleep(2)
        count=ie.checkboxes.length
        puts("PSYS #: " +worksheet.Range("b12").value+ " #{count}" + " item(s) found")
        puts("Test: PSYS # - PASSED")
        ie.image(:index, 2).click
    else
        sleep(2)
        count=ie.checkboxes.length
        puts("PSYS #: " +worksheet.Range("b12").value+ " #{count}" + " item(s) found")
        puts("Test: PSYS # - PASSED")
        ie.button(:name, "method:cancel").click
    end
end 
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click


####################################
########Search Using Asys # as search criteria####
####################################
puts "  "
puts "Asys as Search Criteria"
sleep(2)
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.productNumberType").select("Asys#")
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.productNumber").set(worksheet.Range("b13").value)
ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click
if (ie.pageContainsText("0 items found."))
    count=ie.checkboxes.length
    puts("ASYS #: " +worksheet.Range("b13").value+ " #{count}" + " item(s) found")
    puts("Test: ASYS # - FAILED")
    ie.button(:name, "method:cancel").click
else
    if(ie.pageContainsText("Found PIM Groups"))
        ie.button(:name, "method:executeShowAll").click
        sleep(2)
        count=ie.checkboxes.length
        puts("ASYS #: " +worksheet.Range("b13").value+ " #{count}" + " item(s) found")
        puts("Test: ASYS # - PASSED")
        ie.image(:index, 2).click
    else
        sleep(2)
        count=ie.checkboxes.length
        puts("ASYS #: " +worksheet.Range("b13").value+ " #{count}" + " item(s) found")
        puts("Test: ASYS # - PASSED")
        ie.button(:name, "method:cancel").click
    end
end 
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click

####################################
########Search Using UPC COde as search criteria####
#####################################
puts "  "
puts "UPC Code as Search Criteria"
sleep(2)
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.productNumberType").select("UPC")
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.productNumber").set(worksheet.Range("b14").value)
ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click
if (ie.pageContainsText("0 items found."))
    count=ie.checkboxes.length
    puts("UPC: " +worksheet.Range("b14").value+ " #{count}" + " item(s) found")
    puts("Test: UPC - FAILED")
    ie.button(:name, "method:cancel").click
else
    if(ie.pageContainsText("Found PIM Groups"))
        ie.button(:name, "method:executeShowAll").click
        sleep(2)
        count=ie.checkboxes.length
        puts("UPC: " +worksheet.Range("b14").value+ " #{count}" + " item(s) found")
        puts("Test: UPC - PASSED")
        ie.image(:index, 2).click
    else
        sleep(2)
        count=ie.checkboxes.length
        puts("UPC: " +worksheet.Range("b14").value+ " #{count}" + " item(s) found")
        puts("Test: UPC - PASSED")
        ie.button(:name, "method:cancel").click
    end
end 
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click

####################################
########Search Using USF Prod # as search criteria####
#####################################
puts "  "
puts "USFProd as Search Criteria"
sleep(2)
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.productNumberType").select("USF Prod#")
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.productNumber").set(worksheet.Range("b15").value)
ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click
if (ie.pageContainsText("0 items found."))
    count=ie.checkboxes.length
    puts("USF Prod #: " +worksheet.Range("b15").value+ " #{count}" + " item(s) found")
    puts("Test: USF Prod # - FAILED")
    ie.button(:name, "method:cancel").click
else
    if(ie.pageContainsText("Found PIM Groups"))
        ie.button(:name, "method:executeShowAll").click
        sleep(2)
        count=ie.checkboxes.length
        puts("USF Prod #: " +worksheet.Range("b15").value+ " #{count}" + " item(s) found")
        puts("Test: USF Prod # - PASSED")
        ie.image(:index, 2).click
    else
        sleep(2)
        count=ie.checkboxes.length
        puts("USF Prod #: " +worksheet.Range("b15").value+ " #{count}" + " item(s) found")
        puts("Test: USF Prod # - PASSED")
        ie.button(:name, "method:cancel").click
    end
end 
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click

####################################
########Search Using Prod Master# as search criteria####
#####################################
puts "  "
puts "ProdMaster as Search Criteria"
sleep(2)
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.productNumberType").select("Prod Master#")
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.productNumber").set(worksheet.Range("b16").value)
ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click
if (ie.pageContainsText("0 items found."))
    count=ie.checkboxes.length
    puts("Prod Master #: " +worksheet.Range("b16").value+ " #{count}" + " item(s) found")
    puts("Test: Prod Master # - FAILED")
    ie.button(:name, "method:cancel").click
else
    if(ie.pageContainsText("Found PIM Groups"))
        ie.button(:name, "method:executeShowAll").click
        sleep(2)
        count=ie.checkboxes.length
        puts("Prod Master #: " +worksheet.Range("b16").value+ " #{count}" + " item(s) found")
        puts("Test: Prod Master # - PASSED")
        ie.image(:index, 2).click
    else
        sleep(2)
        count=ie.checkboxes.length
        puts("Prod Master #: " +worksheet.Range("b16").value+ " #{count}" + " item(s) found")
        puts("Test: Prod Master # - PASSED")
        ie.button(:name, "method:cancel").click
    end
end 
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click

####################################
####################################
########Search Using GTIN as search criteria####
#####################################
puts "  "
puts "GTIN as Search Criteria"
sleep(2)
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.productNumberType").select("GTIN")
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.productNumber").set(worksheet.Range("b17").value)
ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click

if (ie.pageContainsText("0 items found."))
    count=ie.checkboxes.length
    puts("GTIN : " +worksheet.Range("b17").value+ " #{count}" + " item(s) found")
    puts("Test: GTIN - FAILED")
    ie.button(:name, "method:cancel").click
else
    if(ie.pageContainsText("Found PIM Groups"))
        ie.button(:name, "method:executeShowAll").click
        sleep(2)
        count=ie.checkboxes.length
        puts("GTIN : " +worksheet.Range("b17").value+ " #{count}" + " item(s) found")
        puts("Test: GTIN - PASSED")
        ie.image(:index, 2).click
    else
        sleep(2)
        count=ie.checkboxes.length
        puts("GTIN : " +worksheet.Range("b17").value+ " #{count}" + " item(s) found")
        puts("Test: GTIN - PASSED")
        ie.button(:name, "method:cancel").click
    end
end 
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click

####################################
#####Search using Mfr  name As Search Criteria###
#####################################
puts "  "
puts "Mfr Name as Search Criteria"  
sleep(2)
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.brandName").set(worksheet.Range("b18").value)
ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click
if (ie.pageContainsText("0 items found."))
    count=ie.checkboxes.length
    puts("Mfr Name: " +worksheet.Range("b18").value+ " #{count}" + " item(s) found")
    puts("Test: Mfr Name - FAILED")
    ie.button(:name, "method:cancel").click
else
    if(ie.pageContainsText("Found PIM Groups"))
        ie.button(:name, "method:executeShowAll").click
        sleep(2)
        count=ie.checkboxes.length
        puts("Mfr Name: " +worksheet.Range("b18").value+ " #{count}" + " item(s) found")
        puts("Test: Mfr Name - PASSED")
        ie.image(:index, 2).click
    else
        sleep(2)
        count=ie.checkboxes.length
        puts("Mfr Name: " +worksheet.Range("b18").value+ " #{count}" + " item(s) found")
        puts("Test: Mfr Name - PASSED")
        ie.button(:name, "method:cancel").click
    end
end 
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click

####################################
#####Search using Brand name As Search Criteria###
#####################################
puts "  "
puts "Brand Name Name As search Criteria"
sleep(2)
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.brandName").set(worksheet.Range("b19").value)
ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click
if (ie.pageContainsText("0 items found."))
    count=ie.checkboxes.length
    puts("Brand Name: " +worksheet.Range("b19").value+ " #{count}" + " item(s) found")
    puts("Test: Brand Name - FAILED")
    ie.button(:name, "method:cancel").click
else
    if(ie.pageContainsText("Found PIM Groups"))
        ie.button(:name, "method:executeShowAll").click
        sleep(2)
        count=ie.checkboxes.length
        puts("Brand Name: " +worksheet.Range("b19").value+ " #{count}" + " item(s) found")
        puts("Test: Brand Name - PASSED")
        ie.image(:index, 2).click
    else
        sleep(2)
        count=ie.checkboxes.length
        puts("Brand Name: " +worksheet.Range("b19").value+ " #{count}" + " item(s) found")
        puts("Test: Brand Name - PASSED")
        ie.button(:name, "method:cancel").click
    end
end 
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click

####################################
#####Search using USF manufacturer # as  As Search Criteria###
#####################################
puts "  "
puts "USF Manufacturer # as Search Criteria"
sleep(2)
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.usfManufacturerNumber").set(worksheet.Range("b20").value)
ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click
if (ie.pageContainsText("0 items found."))
    count=ie.checkboxes.length
    puts("USF Manufacturer #: " +worksheet.Range("b20").value+ " #{count}" + " item(s) found")
    puts("Test: USF Manufacturer # - FAILED")
    ie.button(:name, "method:cancel").click
else
    if(ie.pageContainsText("Found PIM Groups"))
        ie.button(:name, "method:executeShowAll").click
        sleep(2)
        count=ie.checkboxes.length
        puts("USF Manufacturer #: " +worksheet.Range("b20").value+ " #{count}" + " item(s) found")
        puts("Test: USF Manufacturer # - PASSED")
        ie.image(:index, 2).click
    else
        sleep(2)
        count=ie.checkboxes.length
        puts("USF Manufacturer #: " +worksheet.Range("b20").value+ " #{count}" + " item(s) found")
        puts("Test: USF Manufacturer # - PASSED")
        ie.button(:name, "method:cancel").click
    end
end
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click

###########################################
##########Test Brand Type Filter#####################
##########################################
puts "  "
puts "Test: Brand Type Filter"
sleep(1)
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.filterBrand").select("1 - National")
sleep(1)
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.filterBrand").select("2 - Exclusive")
sleep(1)
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.filterBrand").select("3 - Customer")
sleep(1)
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.filterBrand").select("4 - Packer")
sleep(1)
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.filterBrand").select("All Brands")

###################################################
######################Test status Filter################
################################################
puts "  "
puts " test Status Filter Radio Button"
sleep(1)
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.filterStatus").select("Active Only")
sleep(1)
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.filterStatus").select("Inactive Only")
sleep(1)
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.filterStatus").select("All")

#####################################################
###########################test Other Check box################
######################################################
puts "  "
puts "test Status Check Box"
sleep(0.25)
ie.frame("main").frame(:name,"query").checkbox(:value, "KOSHER").set
sleep(0.25)
ie.frame("main").frame(:name,"query").checkbox(:value, "TFF").set
sleep(0.25)
ie.frame("main").frame(:name,"query").checkbox(:value, "CNC").set
sleep(0.25)
ie.frame("main").frame(:name,"query").checkbox(:value, "DATE").set
sleep(0.25)
ie.frame("main").frame(:name,"query").checkbox(:value, "CAB").set
sleep(0.25)
ie.frame("main").frame(:name,"query").checkbox(:value, "HAZARDOUS").set       
sleep(0.25)
ie.frame("main").frame(:name,"query").checkbox(:value, "ORGANIC").set
sleep(0.25)
ie.frame("main").frame(:name,"query").checkbox(:value, "KOSHER").clear
sleep(0.25)
ie.frame("main").frame(:name,"query").checkbox(:value, "TFF").clear
sleep(0.25)
ie.frame("main").frame(:name,"query").checkbox(:value, "CNC").clear
sleep(0.25)
ie.frame("main").frame(:name,"query").checkbox(:value, "DATE").clear
sleep(0.25)
ie.frame("main").frame(:name,"query").checkbox(:value, "CAB").clear
sleep(0.25)
ie.frame("main").frame(:name,"query").checkbox(:value, "HAZARDOUS").clear      
sleep(0.25)
ie.frame("main").frame(:name,"query").checkbox(:value, "ORGANIC").clear
sleep(0.25)
ie.frame("main").frame(:name,"query").checkbox(:value, "ALL").set

#######################################################
#################test Brand Type 1###########################
#######################################################
puts "  "
puts "Test Brand Type 1"
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.filterBrand").select("1 - National")
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.searchKeywords").set(worksheet.Range("b21").value)
ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click
if (ie.pageContainsText("0 items found."))
    count=ie.checkboxes.length
    puts("Brand Type 1: " +worksheet.Range("b21").value+ " #{count}" + " item(s) found")
    puts("Test: Brand Type 1 - FAILED")
    ie.button(:name, "method:cancel").click
else
    if(ie.pageContainsText("Found PIM Groups"))
        ie.button(:name, "method:executeShowAll").click
        sleep(2)
        count=ie.checkboxes.length
        puts("Brand Type 1: " +worksheet.Range("b21").value+ " #{count}" + " item(s) found")
        puts("Test: Brand Type 1 - PASSED")
        ie.image(:index, 2).click
    else
        sleep(2)
        count=ie.checkboxes.length
        puts("Brand Type 1: " +worksheet.Range("b21").value+ " #{count}" + " item(s) found")
        puts("Test: Brand Type 1 - PASSED")
        ie.button(:name, "method:cancel").click
    end
end
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click

#######################################################
#################test Brant d Type 2 ###########################
#######################################################
puts "  "
puts "Test Brand Type 2"
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.filterBrand").select("2 - Exclusive")
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.searchKeywords").set(worksheet.Range("b22").value)
ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click
if (ie.pageContainsText("0 items found."))
    count=ie.checkboxes.length
    puts("Brand Type 2: " +worksheet.Range("b22").value+ " #{count}" + " item(s) found")
    puts("Test: Brand Type 2 - FAILED")
    ie.button(:name, "method:cancel").click
else
    if(ie.pageContainsText("Found PIM Groups"))
        ie.button(:name, "method:executeShowAll").click
        sleep(2)
        count=ie.checkboxes.length
        puts("Brand Type 2: " +worksheet.Range("b22").value+ " #{count}" + " item(s) found")
        puts("Test: Brand Type 2 - PASSED")
        ie.image(:index, 2).click
    else
        sleep(2)
        count=ie.checkboxes.length
        puts("Brand Type 2: " +worksheet.Range("b22").value+ " #{count}" + " item(s) found")
        puts("Test: Brand Type 2 - PASSED")
        ie.button(:name, "method:cancel").click
    end
end
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click

#######################################################
#################test Brant d Type 3 ###########################
#######################################################
puts "  "
puts "Test Brand Type 3"
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.filterBrand").select("3 - Customer")
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.searchKeywords").set(worksheet.Range("b23").value)
ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click
if (ie.pageContainsText("0 items found."))
    count=ie.checkboxes.length
    puts("Brand Type 3: " +worksheet.Range("b23").value+ " #{count}" + " item(s) found")
    puts("Test: Brand Type 3 - FAILED")
    ie.button(:name, "method:cancel").click
else
    if(ie.pageContainsText("Found PIM Groups"))
        ie.button(:name, "method:executeShowAll").click
        sleep(2)
        count=ie.checkboxes.length
        puts("Brand Type 3: " +worksheet.Range("b23").value+ " #{count}" + " item(s) found")
        puts("Test: Brand Type 3 - PASSED")
        ie.image(:index, 2).click
    else
        sleep(2)
        count=ie.checkboxes.length
        puts("Brand Type 3: " +worksheet.Range("b23").value+ " #{count}" + " item(s) found")
        puts("Test: Brand Type 3 - PASSED")
        ie.button(:name, "method:cancel").click
    end
end
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click

#######################################################
#################test Brant d Type 4 ###########################
#######################################################
puts "  "
puts "Test Brand Type 4"
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.filterBrand").select("4 - Packer")
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.searchKeywords").set(worksheet.Range("b24").value)
ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click
if (ie.pageContainsText("0 items found."))
    count=ie.checkboxes.length
    puts("Brand Type 4: " +worksheet.Range("b24").value+ " #{count}" + " item(s) found")
    puts("Test: Brand Type 4 - FAILED")
    ie.button(:name, "method:cancel").click
else
    if(ie.pageContainsText("Found PIM Groups"))
        ie.button(:name, "method:executeShowAll").click
        sleep(2)
        count=ie.checkboxes.length
        puts("Brand Type 4: " +worksheet.Range("b24").value+ " #{count}" + " item(s) found")
        puts("Test: Brand Type 4 - PASSED")
        ie.image(:index, 2).click
    else
        sleep(2)
        count=ie.checkboxes.length
        puts("Brand Type 4: " +worksheet.Range("b24").value+ " #{count}" + " item(s) found")
        puts("Test: Brand Type 4 - PASSED")
        ie.button(:name, "method:cancel").click
    end
end
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click
  
##################################################
####################test Only Active Dropdown ###############
##################################################
puts "  "
puts "Test Status - Active Only"
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.searchKeywords").set(worksheet.Range("b25").value)
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.filterStatus").select("Active Only")
ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click
if (ie.pageContainsText("0 items found."))
    count=ie.checkboxes.length
    puts("Active Only: " +worksheet.Range("b25").value+ " #{count}" + " item(s) found")
    puts("Test: Active Only - FAILED")
    ie.button(:name, "method:cancel").click
else
    if(ie.pageContainsText("Found PIM Groups"))
        ie.button(:name, "method:executeShowAll").click
        sleep(2)
        count=ie.checkboxes.length
        puts("Active Only: " +worksheet.Range("b25").value+ " #{count}" + " item(s) found")
        puts("Test: Active Only - PASSED")
        ie.image(:index, 2).click
    else
        sleep(2)
        count=ie.checkboxes.length
        puts("Active Only: " +worksheet.Range("b25").value+ " #{count}" + " item(s) found")
        puts("Test: Active Only - PASSED")
        ie.button(:name, "method:cancel").click
    end
end
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click

###########################################
####################test Status -Inactive Only###############
##########################################
puts "  "
puts "Test Status -  Inactive Only"
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.searchKeywords").set(worksheet.Range("b26").value)
ie.frame("main").frame(:name,"query").select_list( :name , "searchBean.filterStatus").select("Inactive Only")
ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click
if (ie.pageContainsText("0 items found."))
    count=ie.checkboxes.length
    puts("Inactive Only: " +worksheet.Range("b26").value+ " #{count}" + " item(s) found")
    puts("Test: Inactive Only - FAILED")
    ie.button(:name, "method:cancel").click
else
    if(ie.pageContainsText("Found PIM Groups"))
        ie.button(:name, "method:executeShowAll").click
        sleep(2)
        count=ie.checkboxes.length
        puts("Inactive Only: " +worksheet.Range("b26").value+ " #{count}" + " item(s) found")
        puts("Test: Inactive Only - PASSED")
        ie.image(:index, 2).click
    else
        sleep(2)
        count=ie.checkboxes.length
        puts("Inactive Only: " +worksheet.Range("b26").value+ " #{count}" + " item(s) found")
        puts("Test: Inactive Only - PASSED")
        ie.button(:name, "method:cancel").click
    end
end
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click
###########################################
#################### Kosher ###############
puts "  "
puts "Test kosher Check Box"
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click
sleep(0.25)
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.searchKeywords").set(worksheet.Range("b27").value)
ie.frame("main").frame(:name,"query").checkbox(:value, "KOSHER").set
ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click
if (ie.pageContainsText("0 items found."))
    count=ie.checkboxes.length
    puts("kosher Only: " +worksheet.Range("b27").value+ " #{count}" + " item(s) found")
    puts("Test: Kosher Only - FAILED")
    ie.button(:name, "method:cancel").click
else
    if(ie.pageContainsText("Found PIM Groups"))
        ie.button(:name, "method:executeShowAll").click
        sleep(2)
        count=ie.checkboxes.length
        puts("kosher Only: " +worksheet.Range("b27").value+ " #{count}" + " item(s) found")
        puts("Test: Kosher Only Check Box - PASSED")
        ie.image(:index, 2).click
    else
        sleep(2)
        count=ie.checkboxes.length
        puts("kosher Only: " +worksheet.Range("b27").value+ " #{count}" + " item(s) found")
        puts("Test: Kosher Only Check Box - PASSED")
        ie.button(:name, "method:cancel").click
    end
end
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click

###########################################
#################### TFF ################3
########################################
puts "  "
puts "Test TFF Check Box"
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click
sleep(0.25)
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.searchKeywords").set(worksheet.Range("b28").value)
ie.frame("main").frame(:name,"query").checkbox(:value, "TFF").set
ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click
if (ie.pageContainsText("0 items found."))
    count=ie.checkboxes.length
    puts("TFF Only: " +worksheet.Range("b28").value+ " #{count}" + " item(s) found")
    puts("Test: TFF Only Checkbox - FAILED")
    ie.button(:name, "method:cancel").click
else
    if(ie.pageContainsText("Found PIM Groups"))
        ie.button(:name, "method:executeShowAll").click
        sleep(2)
        count=ie.checkboxes.length
        puts("TFF Only: " +worksheet.Range("b28").value+ " #{count}" + " item(s) found")
        puts("Test: TFF Only Checkbox - PASSED")
        ie.image(:index, 2).click
    else
        sleep(2)
        count=ie.checkboxes.length
        puts("TFF Only: " +worksheet.Range("b28").value+ " #{count}" + " item(s) found")
        puts("Test: TFF Only Checkbox - PASSED")
        ie.button(:name, "method:cancel").click
    end
end
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click
###########################################
####################Hazardous###############
##########################################
puts "  "
puts "Test Hazardous Check Box"
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click
sleep(0.25)
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.searchKeywords").set(worksheet.Range("b29").value)
ie.frame("main").frame(:name,"query").checkbox(:value, "HAZARDOUS").set 
ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click
if (ie.pageContainsText("0 item(s) found."))
    count=ie.checkboxes.length
    puts("Hazardous: " +worksheet.Range("b29").value+ " #{count}" + " item(s) found")
    puts("Test: Hazardousbox - FAILED")
    ie.button(:name, "method:cancel").click
else
    if(ie.pageContainsText("Found PIM Groups"))
        ie.button(:name, "method:executeShowAll").click
        sleep(2)
        count=ie.checkboxes.length
        puts("Hazardous: " +worksheet.Range("b29").value+ " #{count}" + " item(s) found")
        puts("Test: Hazardous Checkbox - PASSED")
        ie.image(:index, 2).click
    else
        sleep(2)
        count=ie.checkboxes.length
        puts("Hazardous: " +worksheet.Range("b29").value+ " #{count}" + " item(s) found")
        puts("Test: Hazardous Checkbox - PASSED")
        ie.button(:name, "method:cancel").click
    end
end
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click
###########################################
####################Hazardous###############
puts "  "
puts "Test CAB Check Box"
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click
sleep(0.25)
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.searchKeywords").set(worksheet.Range("b30").value)
ie.frame("main").frame(:name,"query").checkbox(:value, "CAB").set
ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click
if (ie.pageContainsText("0 items found."))
    count=ie.checkboxes.length
    puts("CAB Only: " +worksheet.Range("b30").value+ " #{count}" + " item(s) found")
    puts("Test: CAB Only Check Box - FAILED")
    ie.button(:name, "method:cancel").click
else
    if(ie.pageContainsText("Found PIM Groups"))
        ie.button(:name, "method:executeShowAll").click
        sleep(2)
        count=ie.checkboxes.length
        puts("CAB Only: " +worksheet.Range("b30").value+ " #{count}" + " item(s) found")
        puts("Test: CAB Only Check Box - PASSED")
        ie.image(:index, 2).click
    else
        sleep(2)
        count=ie.checkboxes.length
        puts("CAB Only: " +worksheet.Range("b30").value+ " #{count}" + " item(s) found")
        puts("Test: CAB Only Check Box - PASSED")
        ie.button(:name, "method:cancel").click
    end
end
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click
###########################################
#################### Organic ###############
puts "  "
puts "Test Organic Check Box"
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click
sleep(0.25)
ie.frame("main").frame(:name,"query").text_field(:name, "searchBean.searchKeywords").set(worksheet.Range("b31").value)
ie.frame("main").frame(:name,"query").checkbox(:value, "ORGANIC").set
ie.frame("main").frame(:name,"query").button(:name, "method:executeSearchQuery").click
if (ie.pageContainsText("0 items found."))
    count=ie.checkboxes.length
    puts("Organic Only : " +worksheet.Range("b31").value+ " #{count}" + " item(s) found")
    puts("Test: Organic Only Checkbox - FAILED")
    ie.button(:name, "method:cancel").click
else
    if(ie.pageContainsText("Found PIM Groups"))
        ie.button(:name, "method:executeShowAll").click
        sleep(2)
        count=ie.checkboxes.length
        puts("Organic Only : " +worksheet.Range("b31").value+ " #{count}" + " item(s) found")
        puts("Test: Organic Only Check box - PASSED")
        ie.image(:index, 2).click
    else
        sleep(2)
        count=ie.checkboxes.length
        puts("Organic Only : " +worksheet.Range("b31").value+ " #{count}" + " item(s) found")
        puts("Test: Organic Only Check box - PASSED")
        ie.button(:name, "method:cancel").click
    end
end      
ie.frame("main").frame(:name,"query").button(:name, "method:executeReset").click
