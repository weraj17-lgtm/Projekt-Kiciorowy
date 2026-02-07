' ===================================================================
' MODUL: GenerateReportModule
' AUTOR: Migracja z Python (generate_report.py)
' DATA: 2026-02-02
' OPIS: Automatyczne generowanie raportow refundacji z formatowaniem
' ===================================================================

Option Explicit

' Glowna procedura uruchamiana przez uzytkownika
Sub GenerateReport()
    Dim ws As Worksheet
    Dim reportDate As String
    Dim reportDateObj As Date
    Dim totalRows As Long
    Dim rowsPerPart As Long
    Dim numParts As Long
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    ' Sprawdz czy mamy aktywny arkusz
    If ActiveWorkbook Is Nothing Then
        MsgBox "Brak otwartego skoroszytu!", vbExclamation
        Exit Sub
    End If
    
    Set ws = ActiveSheet
    
    ' Pobierz date od uzytkownika
    reportDate = InputBox("Podaj date do nazwy pliku (np. 31.01.2026):" & vbCrLf & _
                         "Lub zostaw puste aby pominac date.", _
                         "Data raportu", Format(Date, "DD.MM.YYYY"))
    
    If reportDate = "" Then
        reportDate = vbNullString
    End If
    
    ' Walidacja daty jesli podana
    If reportDate <> vbNullString Then
        On Error Resume Next
        reportDateObj = CDate(reportDate)
        If Err.Number <> 0 Then
            MsgBox "Nieprawidlowy format daty!", vbExclamation
            Exit Sub
        End If
        On Error GoTo ErrorHandler
    End If
    
    ' Przetworz arkusz (czyszczenie i formatowanie)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Call ProcessWorksheet(ws, reportDate)
    
    ' Agregacja danych (nowa funkcjonalnosc zintegrowana z makrem)
    Dim wsResult As Worksheet
    Set wsResult = AggregateWorksheet(ws)
    
    ' Sprawdz czy trzeba dzielic raport
    totalRows = wsResult.UsedRange.Rows.Count - 1
    rowsPerPart = 2000
    
    If totalRows > rowsPerPart Then
        numParts = Application.WorksheetFunction.RoundUp(totalRows / rowsPerPart, 0)
        Call SplitReport(wsResult, reportDate, numParts, rowsPerPart)
        MsgBox "Raport zostal zagregowany i podzielony na " & numParts & " czesci!", vbInformation
    Else
        ' Zapisz pojedynczy plik
        Call SaveReport(wsResult, reportDate, "")
        MsgBox "Raport zostal zagregowany i wygenerowany!", vbInformation
    End If
    
    ' Usun tymczasowy arkusz wynikowy jesli byl stworzony osobno
    ' (w tej wersji AggregateWorksheet tworzy nowy skoroszyt lub arkusz)
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Blad: " & Err.Description, vbCritical
End Sub

' Przetwarza arkusz - wszystkie transformacje
Sub ProcessWorksheet(ws As Worksheet, reportDate As String)
    Dim lastRow As Long
    Dim row As Long
    Dim col_data As Long, col_lp As Long, col_nk As Long
    Dim col_faktura As Long, col_r As Long, col_h As Long, col_t As Long
    Dim headers As Range
    Dim cell As Range
    Dim val As Double
    Dim invoiceVal As String
    Dim cleanedVal As String
    Dim reportDateObj As Date
    Dim highlightRow As Boolean
    Dim statsNK As Long, statsFaktura As Long, statsT As Long
    Dim statsHighValues As Long, statsGreenRows As Long
    
    ' Inicjalizacja statystyk
    statsNK = 0: statsFaktura = 0: statsT = 0
    statsHighValues = 0: statsGreenRows = 0
    
    ' Znajdz indeksy kolumn
    Set headers = ws.Rows(1)
    col_data = FindColumn(headers, "Data raportu")
    col_lp = FindColumn(headers, "L.p.")
    col_nk = FindColumn(headers, "NK")
    col_faktura = FindColumn(headers, "Nr faktury")
    col_r = FindColumn(headers, "Kwota do wyp" & ChrW(322) & "aty")
    col_h = FindColumn(headers, "Stanowisko Kosztowe")
    col_t = 20 ' Kolumna T
    
    If col_data = 0 Or col_lp = 0 Then
        MsgBox "Nie znaleziono wymaganych kolumn!", vbExclamation
        Exit Sub
    End If
    
    ' Przygotuj date
    If reportDate <> vbNullString Then
        reportDateObj = CDate(reportDate)
    End If
    
    lastRow = ws.UsedRange.Rows.Count
    
    ' Przetwarzaj wiersze
    For row = 2 To lastRow
        highlightRow = False
        
        ' 1. Zaktualizuj date raportu
        If reportDate <> vbNullString And col_data > 0 Then
            ws.Cells(row, col_data).Value = reportDateObj
            ws.Cells(row, col_data).NumberFormat = "DD.MM.YYYY"
        End If
        
        ' 2. Wyczysc L.p.
        If col_lp > 0 Then
            ws.Cells(row, col_lp).ClearContents
        End If
        
        ' 3. Przetwarzaj NK
        If col_nk > 0 Then
            On Error Resume Next
            val = CDbl(ws.Cells(row, col_nk).Value)
            If Err.Number = 0 And val > 700 Then
                ws.Cells(row, col_nk).Value = "700"
                ws.Cells(row, col_nk).NumberFormat = "@"
                statsNK = statsNK + 1
            End If
            On Error GoTo 0
        End If
        
        ' 4. Przetwarzaj numer faktury
        If col_faktura > 0 Then
            invoiceVal = CStr(ws.Cells(row, col_faktura).Value)
            If invoiceVal <> "" Then
                cleanedVal = Replace(Replace(Replace(invoiceVal, "\", "/"), "_", "/"), ";", "/")
                If cleanedVal <> invoiceVal Then
                    ws.Cells(row, col_faktura).Value = cleanedVal
                    statsFaktura = statsFaktura + 1
                End If
                
                ' Sprawdz koncowki faktur (00/25, 00/26, 00/2025, 00/2026 itd.)
                ' Uzywamy Trim() i Like *, aby ignorowac spacje i byc odpornym na dlugosc
                Dim checkVal As String
                checkVal = Trim(cleanedVal)
                
                If checkVal Like "*00/25" Or checkVal Like "*00/26" Or _
                   checkVal Like "*00/2025" Or checkVal Like "*00/2026" Then
                    highlightRow = True
                End If
            End If
        End If
        
        ' 5. Przetwarzaj kolumne T
        If col_t > 0 Then
            Dim txtVal As String
            txtVal = CStr(ws.Cells(row, col_t).Value)
            If txtVal <> "" Then
                cleanedVal = CleanNonPolishChars(txtVal)
                If cleanedVal <> txtVal Then
                    ws.Cells(row, col_t).Value = cleanedVal
                    statsT = statsT + 1
                End If
            End If
        End If
        
        ' 6. Wyroznianie kwot >= 500 (czerwony)
        If col_r > 0 Then
            On Error Resume Next
            val = CDbl(ws.Cells(row, col_r).Value)
            If Err.Number = 0 And val >= 500 Then
                ws.Cells(row, col_r).Font.Color = RGB(255, 0, 0)
                statsHighValues = statsHighValues + 1
            End If
            On Error GoTo 0
        End If
        
        ' 7. Format tekstowy dla kolumny H
        If col_h > 0 Then
            ws.Cells(row, col_h).NumberFormat = "@"
        End If
        
        ' 8. Podswietl wiersz na zielono (faktury 00/25-26)
        If highlightRow Then
            ws.Rows(row).Interior.Color = RGB(198, 239, 206)
            statsGreenRows = statsGreenRows + 1
        End If
    Next row
    
    ' Usun kolumny V i W (22 i 23) przed agregacja
    ws.Columns(23).Delete
    ws.Columns(22).Delete
    
    ' Dodatkowo usuń po nazwie jeśli istnieją
    Call DeleteColumnByName(ws, "Skompletow" & ChrW(322) & " - IDX")
    Call DeleteColumnByName(ws, "Skompletow" & ChrW(322))
    
    ' Wyswietl statystyki
    Debug.Print "Statystyki przetwarzania:"
    Debug.Print "  - Poprawiono NK: " & statsNK
    Debug.Print "  - Poprawiono Nr faktur: " & statsFaktura
    Debug.Print "  - Oczyszczono znaki w T: " & statsT
    Debug.Print "  - Wyrozniono kwot >= 500: " & statsHighValues
    Debug.Print "  - Wyrozniono wierszy (zielonych): " & statsGreenRows
End Sub

' Nowa funkcja agregujaca dane (integracja z makrem)
Function AggregateWorksheet(ws As Worksheet) As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dict As Object
    Dim key As String
    Dim newWs As Worksheet
    Dim headers As Range
    
    ' Indeksy kolumn (dynamiczne)
    Dim col_fv As Long, col_kwota As Long, col_konto As Long
    Dim col_osoba As Long, col_kod As Long
    Dim col_imie As Long, col_nazwisko As Long
    
    Set headers = ws.Rows(1)
    col_fv = FindColumn(headers, "Nr faktury")
    col_kwota = FindColumn(headers, "Kwota do wyp" & ChrW(322) & "aty")
    col_konto = FindColumn(headers, "Nr konta")
    col_osoba = FindColumn(headers, "Dane osoby uprawnionej do pobrania refundacji")
    col_kod = FindColumn(headers, "Adres_poczta") ' Odpowiednik N z makra
    col_imie = FindColumn(headers, "Imi" & ChrW(281) & " pacjenta")
    col_nazwisko = FindColumn(headers, "Nazwisko pacjenta")
    col_nazwisko = FindColumn(headers, "Nazwisko pacjenta")
    
    Set dict = CreateObject("Scripting.Dictionary")
    lastRow = ws.UsedRange.Rows.Count
    
    ' 1. Agregacja do slownika
    For i = 2 To lastRow
        ' Klucz: Konto + Osoba + Kod
        key = Trim(CStr(ws.Cells(i, col_konto).Value)) & "|" & _
              Trim(CStr(ws.Cells(i, col_osoba).Value)) & "|" & _
              Trim(CStr(ws.Cells(i, col_kod).Value))
        
        If dict.exists(key) Then
            ' Suma kwoty
            dict(key & "_amt") = dict(key & "_amt") + CDbl(ws.Cells(i, col_kwota).Value)
            
            ' Laczenie faktur (unikalne)
            Dim fv As String
            fv = Trim(CStr(ws.Cells(i, col_fv).Value))
            If InStr(1, dict(key & "_fv"), fv) = 0 Then
                dict(key & "_fv") = dict(key & "_fv") & ", " & fv
            End If
        Else
            dict.Add key, key ' Placeholder
            dict(key & "_amt") = CDbl(ws.Cells(i, col_kwota).Value)
            dict(key & "_fv") = Trim(CStr(ws.Cells(i, col_fv).Value))
            dict(key & "_beneficjent") = ws.Cells(i, col_imie).Value & " " & ws.Cells(i, col_nazwisko).Value
            dict(key & "_konto") = ws.Cells(i, col_konto).Value
        End If
    Next i
    
    ' 2. Tworzenie arkusza wynikowego
    Set newWs = ws.Parent.Sheets.Add
    newWs.Name = "Wynik_" & Format(Now, "hhmmss")
    
    newWs.Cells(1, 1).Value = "Suma Kwoty"
    newWs.Cells(1, 2).Value = "Beneficjent"
    newWs.Cells(1, 3).Value = "Nr Konta"
    newWs.Cells(1, 4).Value = "konto bankowe z ktorego ida platnosci"
    newWs.Cells(1, 5).Value = "Opis nr fv"
    newWs.Cells(1, 6).Value = "Data realizacji"
    
    Dim rowOut As Long
    rowOut = 2
    Dim k As Variant
    For Each k In dict.Keys
        If Not k Like "*_*" Then ' Tylko glowne klucze
            newWs.Cells(rowOut, 1).Value = dict(k & "_amt")
            newWs.Cells(rowOut, 2).Value = dict(k & "_beneficjent")
            
            newWs.Cells(rowOut, 3).NumberFormat = "@"
            newWs.Cells(rowOut, 3).Value = "'" & dict(k & "_konto")
            
            newWs.Cells(rowOut, 4).Value = "" ' Konto bankowe z ktorego ida platnosci ma byc puste
            
            newWs.Cells(rowOut, 5).NumberFormat = "@"
            newWs.Cells(rowOut, 5).Value = "'" & dict(k & "_fv")
            
            newWs.Cells(rowOut, 6).Value = "" ' Data realizacji ma byc pusta
            
            rowOut = rowOut + 1
        End If
    Next k
    
    Set AggregateWorksheet = newWs
    Set dict = Nothing
End Function

' Dzielenie raportu na czesci (kopiuje wszystkie arkusze)
Sub SplitReport(ws As Worksheet, reportDate As String, numParts As Long, rowsPerPart As Long)
    Dim wb As Workbook
    Dim wbPart As Workbook
    Dim wsPart As Worksheet
    Dim partNum As Long
    Dim startRow As Long, endRow As Long
    Dim totalRows As Long
    Dim romanPart As String
    Dim tempPath As String
    Dim fileName As String
    Dim filePath As String
    Dim mainSheetName As String
    
    Set wb = ws.Parent
    totalRows = ws.UsedRange.Rows.Count - 1
    mainSheetName = ws.Name
    
    ' Zapisz tymczasowa kopie calego skoroszytu
    tempPath = wb.Path & "\temp_processed.xlsx"
    Application.DisplayAlerts = False
    wb.SaveCopyAs tempPath
    Application.DisplayAlerts = True
    
    For partNum = 1 To numParts
        ' Otworz tymczasowa kopie (ze wszystkimi arkuszami)
        Set wbPart = Workbooks.Open(tempPath)
        Set wsPart = wbPart.Sheets(mainSheetName)
        
        ' Oblicz zakres wierszy dla tej czesci
        startRow = (partNum - 1) * rowsPerPart + 2
        endRow = Application.WorksheetFunction.Min(partNum * rowsPerPart + 1, totalRows + 1)
        
        ' Usun wiersze spoza zakresu tylko w glownym arkuszu
        If endRow < wsPart.UsedRange.Rows.Count Then
            wsPart.Rows(endRow + 1 & ":" & wsPart.UsedRange.Rows.Count).Delete
        End If
        
        If startRow > 2 Then
            wsPart.Rows("2:" & startRow - 1).Delete
        End If
        
        ' Przygotuj nazwe pliku
        romanPart = GetRoman(partNum)
        If reportDate <> vbNullString Then
            fileName = "Raport Refundacje " & reportDate & " cz. " & romanPart & ".xlsx"
        Else
            fileName = "Raport Refundacje cz. " & romanPart & ".xlsx"
        End If
        filePath = wb.Path & "\" & fileName
        
        ' Zapisz czesc ze wszystkimi arkuszami
        Application.DisplayAlerts = False
        wbPart.SaveAs filePath
        wbPart.Close SaveChanges:=False
        Application.DisplayAlerts = True
        
        Debug.Print "Zapisano: " & fileName & " (wszystkie arkusze zachowane)"
    Next partNum
    
    ' Usun plik tymczasowy
    On Error Resume Next
    Kill tempPath
    On Error GoTo 0
End Sub

' Zapisuje raport do pliku (kopiuje wszystkie arkusze)
Sub SaveReport(ByVal ws As Worksheet, ByVal reportDate As String, ByVal suffix As String)
    Dim wb As Workbook
    Dim fileName As String
    Dim filePath As String
    Dim tempPath As String
    
    Set wb = ws.Parent
    
    If reportDate <> vbNullString Then
        fileName = "Raport Refundacje " & reportDate & suffix & ".xlsx"
    Else
        fileName = "Raport Refundacje" & suffix & ".xlsx"
    End If
    
    filePath = wb.Path & "\" & fileName
    
    ' Zapisz kopie calego skoroszytu ze wszystkimi arkuszami
    Application.DisplayAlerts = False
    wb.SaveCopyAs filePath
    Application.DisplayAlerts = True
    
    Debug.Print "Zapisano: " & fileName & " (wszystkie arkusze zachowane)"
End Sub

' Niezawodna funkcja czyszczaca znaki niepolskie (podejscie whitelist)
Function CleanNonPolishChars(text As String) As String
    Dim i As Long
    Dim char As String
    Dim charCode As Long
    Dim result As String
    
    result = ""
    
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        charCode = AscW(char)
        
        ' Dozwolone: ASCII 32-126 (spacja, litery, cyfry, interpunkcja)
        If charCode >= 32 And charCode <= 126 Then
            result = result & char
        
        ' Polskie znaki (zachowujemy)
        ElseIf IsPolishChar(charCode) Then
            result = result & char
        
        ' Inne znaki - probujemy zamienic na odpowiednik
        Else
            result = result & NormalizeChar(char, charCode)
        End If
    Next i
    
    CleanNonPolishChars = result
End Function

' Sprawdza czy to polski znak
Function IsPolishChar(charCode As Long) As Boolean
    Select Case charCode
        Case 260, 261  ' A, a
            IsPolishChar = True
        Case 262, 263  ' C, c
            IsPolishChar = True
        Case 280, 281  ' E, e
            IsPolishChar = True
        Case 321, 322  ' L, l
            IsPolishChar = True
        Case 323, 324  ' N, n
            IsPolishChar = True
        Case 211, 243  ' O, o
            IsPolishChar = True
        Case 346, 347  ' S, s
            IsPolishChar = True
        Case 377, 378  ' Z, z
            IsPolishChar = True
        Case 379, 380  ' Z, z
            IsPolishChar = True
        Case Else
            IsPolishChar = False
    End Select
End Function

' Normalizuje znak zagraniczny do ASCII
Function NormalizeChar(char As String, charCode As Long) As String
    Select Case charCode
        ' Niemieckie
        Case 228: NormalizeChar = "a"  ' ä
        Case 196: NormalizeChar = "A"  ' Ä
        Case 246: NormalizeChar = "o"  ' ö
        Case 214: NormalizeChar = "O"  ' Ö
        Case 252: NormalizeChar = "u"  ' ü
        Case 220: NormalizeChar = "U"  ' Ü
        Case 223: NormalizeChar = "ss" ' ß
        
        ' Francuskie/ogolne z akcentami
        Case 224, 226: NormalizeChar = "a"  ' à, â
        Case 192, 194: NormalizeChar = "A"  ' À, Â
        Case 233, 232, 234, 235: NormalizeChar = "e"  ' é, è, ê, ë
        Case 201, 200, 202, 203: NormalizeChar = "E"  ' É, È, Ê, Ë
        Case 238, 239: NormalizeChar = "i"  ' î, ï
        Case 206, 207: NormalizeChar = "I"  ' Î, Ï
        Case 244: NormalizeChar = "o"  ' ô
        Case 212: NormalizeChar = "O"  ' Ô
        Case 249, 251: NormalizeChar = "u"  ' ù, û
        Case 217, 219: NormalizeChar = "U"  ' Ù, Û
        Case 231: NormalizeChar = "c"  ' ç
        Case 199: NormalizeChar = "C"  ' Ç
        
        ' Skandynawskie
        Case 229: NormalizeChar = "a"  ' å
        Case 197: NormalizeChar = "A"  ' Å
        Case 230: NormalizeChar = "ae" ' æ
        Case 198: NormalizeChar = "AE" ' Æ
        Case 248: NormalizeChar = "o"  ' ø
        Case 216: NormalizeChar = "O"  ' Ø
        
        ' Hiszpanskie
        Case 241: NormalizeChar = "n"  ' ñ
        Case 209: NormalizeChar = "N"  ' Ñ
        Case 225: NormalizeChar = "a"  ' á
        Case 193: NormalizeChar = "A"  ' Á
        Case 237: NormalizeChar = "i"  ' í
        Case 205: NormalizeChar = "I"  ' Í
        Case 250: NormalizeChar = "u"  ' ú
        Case 218: NormalizeChar = "U"  ' Ú
        
        ' Czeskie/slowackie
        Case 269: NormalizeChar = "c"  ' č
        Case 268: NormalizeChar = "C"  ' Č
        Case 345: NormalizeChar = "r"  ' ř
        Case 344: NormalizeChar = "R"  ' Ř
        Case 353: NormalizeChar = "s"  ' š
        Case 352: NormalizeChar = "S"  ' Š
        Case 382: NormalizeChar = "z"  ' ž
        Case 381: NormalizeChar = "Z"  ' Ž
        Case 357: NormalizeChar = "t"  ' ť
        Case 356: NormalizeChar = "T"  ' Ť
        Case 271: NormalizeChar = "d"  ' ď
        Case 270: NormalizeChar = "D"  ' Ď
        Case 283: NormalizeChar = "e"  ' ě
        Case 282: NormalizeChar = "E"  ' Ě
        Case 367: NormalizeChar = "u"  ' ů
        Case 366: NormalizeChar = "U"  ' Ů
        Case 253: NormalizeChar = "y"  ' ý
        Case 221: NormalizeChar = "Y"  ' Ý
        
        ' Wegierskie
        Case 337: NormalizeChar = "o"  ' ő
        Case 336: NormalizeChar = "O"  ' Ő
        Case 369: NormalizeChar = "u"  ' ű
        Case 368: NormalizeChar = "U"  ' Ű
        
        ' Tureckie
        Case 287: NormalizeChar = "g"  ' ğ
        Case 286: NormalizeChar = "G"  ' Ğ
        Case 305: NormalizeChar = "i"  ' ı
        Case 304: NormalizeChar = "I"  ' İ
        Case 351: NormalizeChar = "s"  ' ş
        Case 350: NormalizeChar = "S"  ' Ş
        
        ' Rumunskie
        Case 259: NormalizeChar = "a"  ' ă
        Case 258: NormalizeChar = "A"  ' Ă
        Case 539: NormalizeChar = "t"  ' ț
        Case 538: NormalizeChar = "T"  ' Ț
        
        ' Inne
        Case 339: NormalizeChar = "oe" ' œ
        Case 338: NormalizeChar = "OE" ' Œ
        Case 255: NormalizeChar = "y"  ' ÿ
        Case 240: NormalizeChar = "d"  ' ð
        Case 208: NormalizeChar = "D"  ' Ð
        Case 254: NormalizeChar = "th" ' þ
        Case 222: NormalizeChar = "TH" ' Þ
        
        ' Nieznany znak - usun
        Case Else
            NormalizeChar = ""
    End Select
End Function

' Konwersja liczby na cyfre rzymska
Function GetRoman(ByVal n As Long) As String
    Select Case n
        Case 1: GetRoman = "I"
        Case 2: GetRoman = "II"
        Case 3: GetRoman = "III"
        Case 4: GetRoman = "IV"
        Case 5: GetRoman = "V"
        Case 6: GetRoman = "VI"
        Case 7: GetRoman = "VII"
        Case 8: GetRoman = "VIII"
        Case 9: GetRoman = "IX"
        Case 10: GetRoman = "X"
        Case 11: GetRoman = "XI"
        Case 12: GetRoman = "XII"
        Case 13: GetRoman = "XIII"
        Case 14: GetRoman = "XIV"
        Case 15: GetRoman = "XV"
        Case 16: GetRoman = "XVI"
        Case 17: GetRoman = "XVII"
        Case 18: GetRoman = "XVIII"
        Case 19: GetRoman = "XIX"
        Case 20: GetRoman = "XX"
        Case Else: GetRoman = CStr(n)
    End Select
End Function

' Znalezienie kolumny po nazwie naglowka
Function FindColumn(headerRow As Range, columnName As String) As Long
    Dim cell As Range
    Dim col As Long
    
    col = 0
    For Each cell In headerRow.Cells
        If Trim(CStr(cell.Value)) = columnName Then
            col = cell.Column
            Exit For
        End If
    Next cell
    
    FindColumn = col
End Function

' Usuwanie kolumny po nazwie
Sub DeleteColumnByName(ws As Worksheet, columnName As String)
    Dim col As Long
    
    col = FindColumn(ws.Rows(1), columnName)
    If col > 0 Then
        ws.Columns(col).Delete
        Debug.Print "  - Usunieto kolumne: " & columnName
    End If
End Sub
