Dim fso, folder, file, targetFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder("c:\Users\Weronika\Desktop\Projekt Kiciorowy")

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

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
Set objWorkbook = objExcel.Workbooks.Open(targetFile)
Set objSheet = objWorkbook.Sheets(1)

found = False
For i = 1 To 100
    val = objSheet.Cells(1, i).Value
    If Not IsEmpty(val) Then
        If InStr(1, val, "za", 1) > 0 Then
            WScript.Echo "Possible match at " & i & ": " & val
            found = True
        End If
    End If
Next

If Not found Then WScript.Echo "No column with 'za' found."

objWorkbook.Close False
objExcel.Quit
