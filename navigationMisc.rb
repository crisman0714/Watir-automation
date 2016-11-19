#requires
require 'watir'

#includes
include Watir

ie = IE.new

ie.goto( 'http://localhost:8080/alpha-swipecard-web/faces/pages/index.jsp' )


ie.text_field( :name, 'j_username' ).set( 'cgmanuel' )
ie.button( :name, 'login' ).click

ie.frame( :name, 'header' ).link( :id, 'form1:Navigation:navPnl:shift' ).click


ie.frame( :name, 'main' ).image( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal:_datePicker_datePickerLink_image' ).click
ie.frame( :name, 'main' ).select_list( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal:_datePicker:monthMenu_list' ).select( 'February' )

ie.frame( :name, 'main' ).image( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal:_datePicker:previousMonthLink_image' ).click
ie.frame( :name, 'main' ).image( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal:_datePicker:previousMonthLink_image' ).click
ie.frame( :name, 'main' ).image( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal:_datePicker:previousMonthLink_image' ).click
ie.frame( :name, 'main' ).image( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal:_datePicker:previousMonthLink_image' ).click

ie.frame( :name, 'main' ).image( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal:_datePicker:nextMonthLink_image' ).click

ie.frame( :name, 'main' ).link( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal:_datePicker_link:6' ).click

ie.frame( :name, 'main' ).image( :id, 'form1:layoutPanel2:groupPanel1:dateToCal:_datePicker_datePickerLink_image' ).click
ie.frame( :name, 'main' ).select_list( :id, 'form1:layoutPanel2:groupPanel1:dateToCal:_datePicker:monthMenu_list' ).select( 'February' )


ie.frame( :name, 'main' ).link( :id, 'form1:layoutPanel2:groupPanel1:dateToCal:_datePicker_link:19' ).click






ie.frame( :name, 'main' ).image( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal:_datePicker_datePickerLink_image' ).click

ie.frame( :name, 'main' ).select_list( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal:_datePicker:monthMenu_list' ).select( 'November' )


ie.frame( :name, 'main' ).select_list( :id, 'form1:layoutPanel2:groupPanel1:dateFromCal:_datePicker:yearMenu_list' ).select( '2008' )


ie.frame( :name, 'main' ).image( :id, 'form1:layoutPanel2:groupPanel1:dateToCal:_datePicker_datePickerLink_image' ).click
ie.frame( :name, 'main' ).select_list( :id, 'form1:layoutPanel2:groupPanel1:dateToCal:_datePicker:monthMenu_list' ).select( 'February' )


ie.frame( :name, 'main' ).link( :id, 'form1:layoutPanel2:groupPanel1:dateToCal:_datePicker_link:19' ).click

ie.frame( :name, 'main' ).button( :id, 'form1:layoutPanel2:searchBtn' ).click

ie.frame( :name, 'main' ).image( :id, 'form1:shiftInOutSummary:_tableActionsBottom:_paginationNextButton:_paginationNextButton_image' ).click

ie.frame( :name, 'main' ).image( :id, 'form1:shiftInOutSummary:_tableActionsBottom:_paginationLastButton:_paginationLastButton_image' ).click

ie.frame( :name, 'main' ).text_field( :id, 'form1:shiftInOutSummary:_tableActionsBottom:_paginationPageField_field' ).set( '27' )
ie.frame( :name, 'main' ).image( :id, 'form1:shiftInOutSummary:_tableActionsBottom:_paginationPrevButton:_paginationPrevButton_image' ).click

ie.frame( :name, 'main' ).image( :id, 'form1:shiftInOutSummary:_tableActionsBottom:_paginationFirstButton:_paginationFirstButton_image' ).click



ie.frame( :name, 'main' ).text_field( :id, 'form1:shiftInOutSummary:_tableActionsBottom:_paginationPageField_field' ).set( '3' )
ie.frame( :name, 'main' ).button( :id, 'form1:shiftInOutSummary:_tableActionsBottom:_paginationSubmitButton' ).click

ie.frame( :name, 'main' ).button( :id, 'form1:shiftInOutSummary:_tableActionsBottom:_paginationSubmitButton' ).click

ie.frame( :name, 'main' ).image( :id, 'form1:shiftInOutSummary:_tableActionsBottom:_paginateButton:_paginateButton_image' ).click
