#requires
require 'watir'

#includes
include Watir

ie = IE.new

ie.goto( 'http://doriath:8092/webcm' )
# DEBUG: Unsupported onfocusout tagname BODY (:index, '7')
# DEBUG: Unsupported action 'click' for 'INPUT'.

# DEBUG: Unsupported action 'click' for 'INPUT'.

ie.text_field( :id, 'username' ).set( 'system' )
ie.button( :name, 'btnSubmit' ).click
# DEBUG: Unsupported onfocusout tagname HTML (:index, '0')
unknown property or method `document'
    HRESULT error code:0x80020006
      Unknown name.
unknown property or method `document'
    HRESULT error code:0x80020006
      Unknown name.
unknown property or method `document'
    HRESULT error code:0x80020006
      Unknown name.
# DEBUG: Unsupported onfocusout tagname FRAMESET (:index, '3')
ie.frame( :name, 'treeframe' ).image( :id, 'webfx-tree-object-7-plus' ).click
# DEBUG: Unsupported onfocusout tagname TD (:index, '14')
ie.frame( :name, 'treeframe' ).image( :id, 'webfx-tree-object-8-image' ).click
# DEBUG: Unsupported onfocusout tagname A (:id, 'webfx-tree-object-8-anchor')
ie.frame( :name, 'treeframe' ).link( :id, 'webfx-menu-object-90' ).click
# DEBUG: Unsupported onfocusout tagname A (:id, 'webfx-menu-object-90')
unknown property or method `document'
    HRESULT error code:0x80020006
      Unknown name.
# DEBUG: Unsupported onfocusout tagname TD (:index, '14')
# DEBUG: Unsupported onfocusout tagname FRAME (:name, 'treeframe')
ie.frame( :name, 'baseframe' ).text_field( :name, 'repositoryName' ).set( 'APIFi
leTest' )
ie.frame( :name, 'baseframe' ).document.all[ '131' ].click
# DEBUG: Unsupported onfocusout tagname HTML (:index, '0')
unknown property or method `document'
    HRESULT error code:0x80020006
      Unknown name.
# DEBUG: Unsupported onfocusout tagname BODY (:index, '11')
# DEBUG: Unsupported action 'click' for 'INPUT'.

ie.frame( :name, 'baseframe' ).checkbox( :name, 'relationColMap(1011385)' ).set
# DEBUG: Unsupported action 'click' for 'INPUT'.

ie.frame( :name, 'baseframe' ).checkbox( :name, 'relationColMap(1011164)' ).set
ie.frame( :name, 'baseframe' ).document.all[ '3854' ].click
# DEBUG: Unsupported onfocusout tagname HTML (:index, '0')
unknown property or method `document'
    HRESULT error code:0x80020006
      Unknown name.
# DEBUG: Unsupported onfocusout tagname BODY (:index, '11')
# DEBUG: Unsupported onfocusout tagname FRAME (:name, 'baseframe')
# DEBUG: Unsupported onfocusout tagname HTML (:index, '0')
unknown property or method `document'
    HRESULT error code:0x80020006
      Unknown name.
# DEBUG: Unsupported onfocusout tagname BODY (:index, '9')
ie.frame( :name, 'baseframe' ).text_field( :name, 'repositoryName' ).set( 'APIFi
leTest' )
# DEBUG: Unsupported onfocusout tagname FRAME (:name, 'baseframe')
