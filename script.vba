Option Explicit
Private Sub copyFromExternalCosts()
    'On Error GoTo ErrHandler
    Dim fileName As String, sheet As Worksheet
    Dim fd As Office.FileDialog

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .AllowMultiSelect = False
        .Title = "Please select the file."
        .Filters.Clear
        .Filters.Add "Excel 2010", "*.xlsx"
    
        If .Show = True Then
            fileName = Dir(.SelectedItems(1))
        End If
    End With

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    If fileName <> "" Then
        ReadDataFromFile (fileName)
        
        MsgBox fileName & " - файл обработан"
    Else
        MsgBox "Не выбран файл"
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub ReadDataFromFile(ByVal filePath As String)
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    Dim src As Workbook
    
    ' OPEN THE SOURCE EXCEL WORKBOOK IN "READ ONLY MODE".
    Set src = Workbooks.Open(filePath, True, True)
    
    ' GET THE TOTAL ROWS FROM THE SOURCE WORKBOOK.
    Dim sheetCostsExternal As Worksheet
    Dim sheetCostsInternal As Worksheet
    
    Set sheetCostsExternal = src.Worksheets("смета")
    Set sheetCostsInternal = getOrCreateSheet("смета")
    
    copyValue sheetCostsExternal.Range("E3"), sheetCostsInternal.Range("E3")
    copyValue sheetCostsExternal.Range("E4"), sheetCostsInternal.Range("E4")
    copyValue sheetCostsExternal.Range("E5"), sheetCostsInternal.Range("E5")
    copyValue sheetCostsExternal.Range("E6"), sheetCostsInternal.Range("E6")
    copyValue sheetCostsExternal.Range("E7"), sheetCostsInternal.Range("E7")
    copyValue sheetCostsExternal.Range("E8"), sheetCostsInternal.Range("E8")
    

    ' Values
    copyValue sheetCostsExternal.Range("B11", "C12"), sheetCostsInternal.Range("B11", "C12")
    
    copyValue sheetCostsExternal.Range("B15", "C43"), sheetCostsInternal.Range("B15", "C43")
    
    ' Company
    copyValue sheetCostsExternal.Range("A49", "A58"), sheetCostsInternal.Range("A49", "A58")
    
    ' Main income
    copyValue sheetCostsExternal.Range("E49", "E58"), sheetCostsInternal.Range("E49", "E58")
    
    ' Lecturer sums
    copyValue sheetCostsExternal.Range("F49", "F58"), sheetCostsInternal.Range("F49", "F58")

    ' Fees
    copyValue sheetCostsExternal.Range("H49", "H58"), sheetCostsInternal.Range("H49", "H58")
    
    ' Legal entities
    copyValue sheetCostsExternal.Range("J49", "J58"), sheetCostsInternal.Range("J49", "J58")
    
    ' CLOSE THE SOURCE FILE.
    src.Close False             ' FALSE - DON'T SAVE THE SOURCE FILE.
    Set src = Nothing
    
ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    'Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Sub copyFromCostsToTable()

    copyFromCostsToTech
    copyFromTechToTable

End Sub

Sub copyFromCostsToTech()
    MsgBox "Start copyFromCostsToTech"
    Dim sheetTech As Worksheet
    Dim sheetCosts As Worksheet
    Dim sheetTable As Worksheet
    
    Set sheetTech = getOrCreateSheet("техн")
    Set sheetCosts = getOrCreateSheet("смета")
    Set sheetTable = getOrCreateSheet("таблица")
    
    Dim rngId As Range
    Dim rngTableStart As Range
    Dim rngTableEnd As Range
    
    Set rngId = sheetCosts.Range("J3")
    Set rngTableStart = sheetTech.Range("B3")
    Set rngTableEnd = sheetTech.Range("B4")
    
    Dim vId As String
    Dim vRowIndex As Integer
    If rngId.Value <> "" Then
        vId = CStr(rngId.Value)
        
        Set rngId = sheetTech.Range("A" & rngTableStart.Value, "A" & rngTableEnd.Value).Find(vId)
        If rngId Is Nothing Then
            rngTableEnd.Value = rngTableEnd.Value + 1
            vRowIndex = CInt(rngTableEnd.Value)
            sheetTech.Range("A" & vRowIndex).Value = vId
        Else
            vRowIndex = rngId.Row
        End If
    Else
        vId = getNextId()
        rngId.Value = vId
        rngTableEnd.Value = rngTableEnd.Value + 1
        vRowIndex = CInt(rngTableEnd.Value)
        sheetTech.Range("A" & vRowIndex).Value = vId
    End If
        
    Set rngId = sheetTech.Range("E1")
    rngId.Value = vId
    
    'Exit Sub
    
    
    ' Headers
    'copyValue sheetCosts.Range("A3"), sheetTech.Range("B1")
    'copyValue sheetCosts.Range("A4"), sheetTech.Range("C1")
    'copyValue sheetCosts.Range("A5"), sheetTech.Range("D1")
    'copyValue sheetCosts.Range("A6"), sheetTech.Range("E1")
    'copyValue sheetCosts.Range("A7"), sheetTech.Range("F1")
    'copyValue sheetCosts.Range("A8"), sheetTech.Range("G1")
    
    ' Values
    copyValue sheetCosts.Range("E3"), sheetTech.Range("B" & vRowIndex)
    copyValue sheetCosts.Range("E4"), sheetTech.Range("C" & vRowIndex)
    copyValue sheetCosts.Range("E5"), sheetTech.Range("D" & vRowIndex)
    copyValue sheetCosts.Range("E6"), sheetTech.Range("E" & vRowIndex)
    copyValue sheetCosts.Range("E7"), sheetTech.Range("F" & vRowIndex)
    copyValue sheetCosts.Range("E8"), sheetTech.Range("G" & vRowIndex)
    
    
    ' Headers
    'copyValue sheetCosts.Range("A11"), sheetTech.Cells(1, 8)
    'copyValue sheetCosts.Range("A12"), sheetTech.Cells(1, 9)
    'copyValue sheetCosts.Range("A15"), sheetTech.Cells(1, 10)
    'copyValue sheetCosts.Range("A16"), sheetTech.Cells(1, 11)
    'copyValue sheetCosts.Range("A17"), sheetTech.Cells(1, 12)
    'copyValue sheetCosts.Range("A18"), sheetTech.Cells(1, 13)
    'copyValue sheetCosts.Range("A19"), sheetTech.Cells(1, 14)
    'copyValue sheetCosts.Range("A20"), sheetTech.Cells(1, 15)
    'copyValue sheetCosts.Range("A21"), sheetTech.Cells(1, 16)
    'copyValue sheetCosts.Range("A22"), sheetTech.Cells(1, 17)
    'copyValue sheetCosts.Range("A23"), sheetTech.Cells(1, 18)
    'copyValue sheetCosts.Range("A24"), sheetTech.Cells(1, 19)
    'copyValue sheetCosts.Range("A25"), sheetTech.Cells(1, 20)
    'copyValue sheetCosts.Range("A26"), sheetTech.Cells(1, 21)
    'copyValue sheetCosts.Range("A27"), sheetTech.Cells(1, 22)
    'copyValue sheetCosts.Range("A28"), sheetTech.Cells(1, 23)
    'copyValue sheetCosts.Range("A29"), sheetTech.Cells(1, 24)
    'copyValue sheetCosts.Range("A30"), sheetTech.Cells(1, 25)
    'copyValue sheetCosts.Range("A31"), sheetTech.Cells(1, 26)
    'copyValue sheetCosts.Range("A32"), sheetTech.Cells(1, 27)
    'copyValue sheetCosts.Range("A33"), sheetTech.Cells(1, 28)
    'copyValue sheetCosts.Range("A34"), sheetTech.Cells(1, 29)
    'copyValue sheetCosts.Range("A35"), sheetTech.Cells(1, 30)
    'copyValue sheetCosts.Range("A36"), sheetTech.Cells(1, 31)
    'copyValue sheetCosts.Range("A37"), sheetTech.Cells(1, 32)
    'copyValue sheetCosts.Range("A38"), sheetTech.Cells(1, 33)
    'copyValue sheetCosts.Range("A39"), sheetTech.Cells(1, 34)
    'copyValue sheetCosts.Range("A40"), sheetTech.Cells(1, 35)
    'copyValue sheetCosts.Range("A41"), sheetTech.Cells(1, 36)
    'copyValue sheetCosts.Range("A42"), sheetTech.Cells(1, 37)
    'copyValue sheetCosts.Range("A43"), sheetTech.Cells(1, 38)

    ' Values
    Dim vStartColIndex As Integer
    vStartColIndex = 8
    For Each cell In sheetCosts.Range("B11", "B12").Cells
        sheetTech.Cells(vRowIndex, vStartColIndex).Value = cell.Offset(, 1).Value & "::" & cell.Value
        vStartColIndex = vStartColIndex + 1
    Next
    
    For Each cell In sheetCosts.Range("B15", "B43").Cells
        sheetTech.Cells(vRowIndex, vStartColIndex).Value = cell.Offset(, 1).Value & "::" & cell.Value
        vStartColIndex = vStartColIndex + 1
    Next
    
    
    ' Company
    For Each cell In sheetCosts.Range("A49", "A58").Cells
        'sheetTech.Cells(1, vStartColIndex).Value = "Компания-участник"
        sheetTech.Cells(vRowIndex, vStartColIndex).Value = cell.Value
        vStartColIndex = vStartColIndex + 1
    Next
    
    
    ' Main income
    For Each cell In sheetCosts.Range("E49", "E58").Cells
        'sheetTech.Cells(1, vStartColIndex).Value = "Основной приход"
        sheetTech.Cells(vRowIndex, vStartColIndex).Value = cell.Value
        vStartColIndex = vStartColIndex + 1
    Next
    
    
    ' Lecturer sums
    For Each cell In sheetCosts.Range("F49", "F58").Cells
        'sheetTech.Cells(1, vStartColIndex).Value = "Лекторские|Сумма"
        sheetTech.Cells(vRowIndex, vStartColIndex).Value = cell.Value
        vStartColIndex = vStartColIndex + 1
    Next
    

    ' Fees
    For Each cell In sheetCosts.Range("H49", "H58").Cells
        'sheetTech.Cells(1, vStartColIndex).Value = "Комиссия"
        sheetTech.Cells(vRowIndex, vStartColIndex).Value = cell.Value
        vStartColIndex = vStartColIndex + 1
    Next
    
    
    ' Legal entities
    For Each cell In sheetCosts.Range("J49", "J58").Cells
        'sheetTech.Cells(1, vStartColIndex).Value = "Юрлицо"
        sheetTech.Cells(vRowIndex, vStartColIndex).Value = cell.Value
        vStartColIndex = vStartColIndex + 1
    Next
    
    'MsgBox vStartColIndex
    'MsgBox vRowIndex
End Sub

Sub copyFromTechToTable()
    MsgBox "Start copyFromTechToTable"
    Dim sheetTech As Worksheet
    Dim sheetCosts As Worksheet
    Dim sheetTable As Worksheet
    
    Set sheetTech = getOrCreateSheet("техн")
    Set sheetCosts = getOrCreateSheet("смета")
    Set sheetTable = getOrCreateSheet("таблица")
    
    Dim rngId As Range
    Dim rngTableStart As Range
    Dim rngTableEnd As Range
    Set rngId = sheetTech.Range("E1")
    Set rngTableStart = sheetTech.Range("B3")
    Set rngTableEnd = sheetTech.Range("B4")
    
    Dim vId As String
    Dim vRowIndexFrom As Integer
    Dim vRowIndexTo As Integer
    If rngId.Value <> "" Then
        vId = CStr(rngId.Value)
        
        Set rngId = sheetTech.Range("A" & rngTableStart.Value, "A" & rngTableEnd.Value).Find(vId)
        If rngId Is Nothing Then
            Err.Raise -20001, "Копирование из Техн", "Не найден текущий идентификатор"
        Else
            vRowIndexFrom = rngId.Row
        End If
    End If
    
    Set rngId = sheetTable.Range("A2")
    If rngId.Value = "" Then
        rngId.Value = vId
        vRowIndexTo = rngId.Row
    Else
        Set rngId = sheetTable.Range("A1", sheetTable.Range("A1").End(xlDown)).Find(vId)
        If rngId Is Nothing Then
            Set rngId = sheetTable.Range("A1").End(xlDown).Offset(1, 0)
            rngId.Value = vId
            vRowIndexTo = rngId.Row
        Else
            vRowIndexTo = rngId.Row
        End If
        
    End If
    
    
    ' Headers
    'copyValue sheetTech.Range("B1", "G1"), sheetTable.Range("B1", "G1")
    
    ' Values
    copyValue sheetTech.Range("B" & vRowIndexFrom, "G" & vRowIndexFrom), sheetTable.Range("B" & vRowIndexTo, "G" & vRowIndexTo)
    
    Dim sumCash As Double
    Dim sumNonCash As Double
    
    sumCash = 0#
    sumNonCash = 0#
    
    For Each cell In sheetTech.Range("H" & vRowIndexFrom, "AL" & vRowIndexFrom).Cells
        If cell.Value <> "" Then
            vals = Split(cell.Value, "::")
            
            If vals(0) = "БН" Then
                sumNonCash = sumNonCash + CDbl(vals(1))
            ElseIf vals(0) = "Н" Then
                sumCash = sumCash + CDbl(vals(1))
            Else
                Err.Raise -20001, "Копирование Техн в Таблица", "Значение отличное от БН и Н для сумм сметы"
            End If
            
        End If
    Next
    
    'sheetTable.Range("H1").Value = "Расходы б/н"
    sheetTable.Range("H" & vRowIndexTo).Value = sumNonCash
    
    'sheetTable.Range("I1").Value = "Расходы н"
    sheetTable.Range("I" & vRowIndexTo).Value = sumCash
    
    
    
    Dim companies As String
    
    companies = ""
    
    For Each cell In sheetTech.Range("AM" & vRowIndexFrom, "AV" & vRowIndexFrom).Cells
        If cell.Value <> "" Then
            If companies = "" Then
                companies = cell.Value
            Else
                companies = companies & ", " & cell.Value
            End If
            
        End If
    Next
    
    'sheetTable.Range("J1").Value = "Название компании"
    sheetTable.Range("J" & vRowIndexTo).Value = companies
    
        
    Dim mainIncome As Double
    mainIncome = 0#
    
    For Each cell In sheetTech.Range("AW" & vRowIndexFrom, "BF" & vRowIndexFrom).Cells
        If cell.Value <> "" Then
            mainIncome = mainIncome + CDbl(cell.Value)
            
        End If
    Next
    
    'sheetTable.Range("K1").Value = "Основной приход"
    sheetTable.Range("K" & vRowIndexTo).Value = mainIncome
    
    
    Dim sumLecturers As Double
    sumLecturers = 0#
    
    For Each cell In sheetTech.Range("BG" & vRowIndexFrom, "BP" & vRowIndexFrom).Cells
        If cell.Value <> "" Then
            sumLecturers = sumLecturers + CDbl(cell.Value)
            
        End If
    Next
    
    'sheetTable.Range("L1").Value = "Лекторские"
    sheetTable.Range("L" & vRowIndexTo).Value = sumLecturers
    
    
    Dim fees As Double
    fees = 0#
    
    For Each cell In sheetTech.Range("BQ" & vRowIndexFrom, "BZ" & vRowIndexFrom).Cells
        If cell.Value <> "" Then
            fees = fees + CDbl(cell.Value)
            
        End If
    Next
    
    'sheetTable.Range("M1").Value = "Комиссия"
    sheetTable.Range("M" & vRowIndexTo).Value = fees
    
    
    Dim legalEntities As String
    
    legalEntities = ""
    
    For Each cell In sheetTech.Range("CA" & vRowIndexFrom, "CJ" & vRowIndexFrom).Cells
        If cell.Value <> "" Then
            If legalEntities = "" Then
                legalEntities = cell.Value
            Else
                legalEntities = legalEntities & ", " & cell.Value
            End If
            
        End If
    Next
    
    'sheetTable.Range("N1").Value = "Юрлица"
    sheetTable.Range("N" & vRowIndexTo).Value = legalEntities
    
End Sub


Sub copyValue(ByRef pFrom As Range, ByRef pTo As Range)

   pTo.Value = pFrom.Value

End Sub

Function getOrCreateSheet(ByRef pSheetName As String) As Worksheet
    On Error GoTo ErrorHandler

    Set getOrCreateSheet = Worksheets(pSheetName)
    Exit Function
    
ErrorHandler:
    Set getOrCreateSheet = Worksheets.Add
    getOrCreateSheet.Name = pSheetName
    getOrCreateSheet.Activate
    Exit Function
End Function

Function getNextId() As String
    Dim sheetTech As Worksheet
    Set sheetTech = getOrCreateSheet("техн")
    
    sheetTech.Range("B2").Value = sheetTech.Range("B2").Value + 1
    getNextId = sheetTech.Range("B1").Value & "_" & sheetTech.Range("B2").Value
    
    MsgBox getNextId
End Function

Function getLastId() As String
    Dim sheetTech As Worksheet
    Set sheetTech = getOrCreateSheet("техн")
    
    getLastId = sheetTech.Range("B1").Value & "_" & sheetTech.Range("B2").Value
    
    MsgBox getLastId
End Function

