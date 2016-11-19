#requires
require 'watir'

#includes
include Watir

ie = IE.new

ie.goto( 'http://google.com/' )
# DEBUG: Unsupported onfocusout tagname BODY (:index, '7')
ie.goto( 'http://www.google.com.ph/search?hl=tl&q=watir&meta=&aq=f&oq=' )
# DEBUG: Unsupported onfocusout tagname HTML (:index, '1')
# DEBUG: Unsupported onfocusout tagname BODY (:id, 'gsr')
# DEBUG: Unsupported onclick tagname EM (:index, '75')
# DEBUG: Unsupported onfocusout tagname HTML (:index, '2')

ie.goto( 'http://wtr.rubyforge.org/robots.txt#resize_iframe%26remote_iframe_0%26126%261$' )
# DEBUG: Unsupported onfocusout tagname BODY (:index, '14')