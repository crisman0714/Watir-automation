#open spreadsheet
excel = WIN32OLE::new('excel.Application')
workbook = excel.Workbooks.Add
worksheet = workbook.Worksheets(1)
worksheet.SaveAs("spreadsheet.xls")
#Log results
worksheet.range("a1").value = executionEnvironment
worksheet.range("b1").value = "Acceptable Screen1 time"
worksheet.range("c1").value = acceptableScreen1.to_s
worksheet.range("d1").value = "Actual Screen1 time"
worksheet.range("e1").value = actualScreen1.to_s
worksheet.range("f1").value = resultValue
#
# Etcetera...assume the above happens 4 times, for 4 screens...
#
#Format workbook columns
worksheet.range("b1:b4").Interior['ColorIndex'] = 36 #pale yellow
worksheet.columns("b:b").AutoFit
#close the workbook
workbook.save
workbook.close
excel.Quit
