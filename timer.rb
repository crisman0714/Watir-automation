#includes
require 'watir.rb'  # the controller

#variables
testSite = 'http://www.BookPool.com'
executionEnvironment = 'Test'  #This could be read from a configuration file at runtime
beginTime = 0
endTime = 0

#open spreadsheet
timeSpreadsheet = File.new( Time.now.strftime("%d-%b-%y") + ".csv", "a")  #Note this creates a new file every day...

#open the IE browser to http://www.BookPool.com
$ie = IE.new
beginTime = Time.now
$ie.goto(testSite)
endTime = Time.now

#Log the time for the Home Page to load
timeSpreadsheet.puts executionEnvironment + ",Home Page," + (endTime - beginTime).to_s

puts 'Step 2: enter "Ruby: in the search text field under Simple Search'
$ie.textField(:name, "qs").set("Ruby") # qs is the name of the search field

puts 'Step 3: submit the search form.'
beginTime = Time.now
$ie.form(:index, "1").submit #submitting first form found in the page.
endTime = Time.now

#Log the time for the search
timeSpreadsheet.puts executionEnvironment + ",Time to execute search," + (endTime - beginTime).to_s

#All 7 results for Ruby:
puts 'Actual Result: Check that the "All 7 results for Ruby:" message actually appears'
a = $ie.pageContainsText("All 7 results for Ruby:")
if !a
  scriptLog.puts "Test Failed! Could not find test string: 'All 7 results for Ruby:'"
else
  scriptLog.puts "Test Passed. Found the test string: 'All 7 results for Ruby:'"
end

puts 'Close the browser'
$ie.close()

#close the file
timeSpreadsheet.close
