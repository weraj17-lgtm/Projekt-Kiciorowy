Dim fso, folder, file, targetFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder("c:\Users\Weronika\Desktop\Projekt Kiciorowy")

For Each file In folder.Files
    If InStr(file.Name, "wz") > 0 And LCase(fso.GetExtensionName(file.Name)) = "xlsx" Then
        targetFile = file.Path
        Exit For
    End If
Next

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
Set objWorkbook = objExcel.Workbooks.Open(targetFile)
Set objSheet = objWorkbook.Sheets(1)

WScript.Echo "Extended Headers:"
For i = 1 To 50
    header = objSheet.Cells(1, i).Value
    If Not IsEmpty(header) Then
        WScript.Echo i & ": " & header
    End If
Next

objWorkbook.Close False
objExcel.Quit
