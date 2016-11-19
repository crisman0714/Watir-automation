require 'watir'   # the controller
require 'win32ole'
include Watir

require 'test/unit' 
require 'test/unit/ui/console/testrunner'
require 'example_logger1'



class TC_MAINTENANCE_REQUESTs < Test::Unit::TestCase
      
   
# List of .rbTest file 
require 'SCloginTest_Negative'
require 'SCloginTest_Positive'
require 'PowerUserSearchSwInOutSingResult'
require 'PowerUserSearchSwInOutMultiResult'
require 'NormalUserSearchSwInOutSingResult1'
require 'NormalUserSearchSwInOutMulResult'
require 'AdminUserSearchMultipleResult'
require 'AdminUserSearchSingleResult'
#require 'misc1'
#require 'misc'


end
