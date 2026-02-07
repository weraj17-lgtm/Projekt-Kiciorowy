Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("c:\Users\Weronika\Desktop\Projekt Kiciorowy\wzór.xlsx")
Set objSheet = objWorkbook.Sheets(1)
WScript.Echo "Columns found:"
For i = 1 To objSheet.UsedRange.Columns.Count
    WScript.Echo i & ": " & objSheet.Cells(1, i).Value
Next
objWorkbook.Close
objExcel.Quit
