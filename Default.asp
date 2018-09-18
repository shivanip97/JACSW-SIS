set objExcel = CreateObject("Excel.Application")
objExcel.Application.DisplayAlerts = False
set objWorkbook=objExcel.workbooks.add()
objExcel.cells(1,1).value = "Test value"
objExcel.cells(1,2).value = "Test data"
objWorkbook.Saveas "c:\testXLS.xls"
objWorkbook.Close
objExcel.workbooks.close
objExcel.quit