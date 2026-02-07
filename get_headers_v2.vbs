Dim fso, folder, file
Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder("c:\Users\Weronika\Desktop\Projekt Kiciorowy")

Dim targetFile
For Each file In folder.Files
    If InStr(file.Name, "wz") > 0 And LCase(fso.GetExtensionName(file.Name)) = "xlsx" Then
        targetFile = file.Path
        Exit For
    End If
Next

If targetFile = "" Then
    WScript.Echo "File not found."
    WScript.Quit
End If

WScript.Echo "Reading: " & targetFile

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
Set objWorkbook = objExcel.Workbooks.Open(targetFile)
Set objSheet = objWorkbook.Sheets(1)

WScript.Echo "Headers:"
For i = 1 To 30
    WScript.Echo i & ": " & objSheet.Cells(1, i).Value
Next

objWorkbook.Close False
objExcel.Quit
