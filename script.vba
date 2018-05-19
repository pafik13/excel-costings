Sub copyFromCostsToTech()
    MsgBox "Start copyFromCostsToTech"
    Dim sheetTech As Worksheet
    Dim sheetCosts As Worksheet
    Dim sheetTable As Worksheet
    
    Set sheetTech = getOrCreateSheet("òåõí")
    Set sheetCosts = getOrCreateSheet("ñìåòà")
    Set sheetTable = getOrCreateSheet("òàáëèöà")
    
    'Dim rngFrom As Range
    'Dim rngTo As Range
    
    'Set rngFrom = sheetÑosts.Range("A3", "A8")
    'Set rngTo = sheetTech.Range("B1", "I1")
    
    'rngTo.Value2 = rngFrom.Value2
    
    ' Headers
    copyValue sheetCosts.Range("A3"), sheetTech.Range("B1")
    copyValue sheetCosts.Range("A4"), sheetTech.Range("C1")
    copyValue sheetCosts.Range("A5"), sheetTech.Range("D1")
    copyValue sheetCosts.Range("A6"), sheetTech.Range("E1")
    copyValue sheetCosts.Range("A7"), sheetTech.Range("F1")
    copyValue sheetCosts.Range("A8"), sheetTech.Range("G1")
    
    ' Values
    copyValue sheetCosts.Range("E3"), sheetTech.Range("B2")
    copyValue sheetCosts.Range("E4"), sheetTech.Range("C2")
    copyValue sheetCosts.Range("E5"), sheetTech.Range("D2")
    copyValue sheetCosts.Range("E6"), sheetTech.Range("E2")
    copyValue sheetCosts.Range("E7"), sheetTech.Range("F2")
    copyValue sheetCosts.Range("E8"), sheetTech.Range("G2")
    
    
    ' Headers
    copyValue sheetCosts.Range("A11"), sheetTech.Cells(1, 8)
    copyValue sheetCosts.Range("A12"), sheetTech.Cells(1, 9)
    copyValue sheetCosts.Range("A15"), sheetTech.Cells(1, 10)
    copyValue sheetCosts.Range("A16"), sheetTech.Cells(1, 11)
    copyValue sheetCosts.Range("A17"), sheetTech.Cells(1, 12)
    copyValue sheetCosts.Range("A18"), sheetTech.Cells(1, 13)
    copyValue sheetCosts.Range("A19"), sheetTech.Cells(1, 14)
    copyValue sheetCosts.Range("A20"), sheetTech.Cells(1, 15)
    copyValue sheetCosts.Range("A21"), sheetTech.Cells(1, 16)
    copyValue sheetCosts.Range("A22"), sheetTech.Cells(1, 17)
    copyValue sheetCosts.Range("A23"), sheetTech.Cells(1, 18)
    copyValue sheetCosts.Range("A24"), sheetTech.Cells(1, 19)
    copyValue sheetCosts.Range("A25"), sheetTech.Cells(1, 20)
    copyValue sheetCosts.Range("A26"), sheetTech.Cells(1, 21)
    copyValue sheetCosts.Range("A27"), sheetTech.Cells(1, 22)
    copyValue sheetCosts.Range("A28"), sheetTech.Cells(1, 23)
    copyValue sheetCosts.Range("A29"), sheetTech.Cells(1, 24)
    copyValue sheetCosts.Range("A30"), sheetTech.Cells(1, 25)
    copyValue sheetCosts.Range("A31"), sheetTech.Cells(1, 26)
    copyValue sheetCosts.Range("A32"), sheetTech.Cells(1, 27)
    copyValue sheetCosts.Range("A33"), sheetTech.Cells(1, 28)
    copyValue sheetCosts.Range("A34"), sheetTech.Cells(1, 29)
    copyValue sheetCosts.Range("A35"), sheetTech.Cells(1, 30)
    copyValue sheetCosts.Range("A36"), sheetTech.Cells(1, 31)
    copyValue sheetCosts.Range("A37"), sheetTech.Cells(1, 32)
    copyValue sheetCosts.Range("A38"), sheetTech.Cells(1, 33)
    copyValue sheetCosts.Range("A39"), sheetTech.Cells(1, 34)
    copyValue sheetCosts.Range("A40"), sheetTech.Cells(1, 35)
    copyValue sheetCosts.Range("A41"), sheetTech.Cells(1, 36)
    copyValue sheetCosts.Range("A42"), sheetTech.Cells(1, 37)
    copyValue sheetCosts.Range("A43"), sheetTech.Cells(1, 38)

    ' Values
    Dim vStartColIndex As Integer
    vStartColIndex = 8
    For Each cell In sheetCosts.Range("B11", "B12").Cells
        sheetTech.Cells(2, vStartColIndex).Value = cell.Offset(, 1).Value & "::" & cell.Value
        vStartColIndex = vStartColIndex + 1
    Next
    
    For Each cell In sheetCosts.Range("B15", "B43").Cells
        sheetTech.Cells(2, vStartColIndex).Value = cell.Offset(, 1).Value & "::" & cell.Value
        vStartColIndex = vStartColIndex + 1
    Next
    
    
    ' Company
    For Each cell In sheetCosts.Range("A49", "A58").Cells
        sheetTech.Cells(1, vStartColIndex).Value = "Êîìïàíèÿ-ó÷àñòíèê"
        sheetTech.Cells(2, vStartColIndex).Value = cell.Value
        vStartColIndex = vStartColIndex + 1
    Next
    
    
    ' Main income
    For Each cell In sheetCosts.Range("E49", "E58").Cells
        sheetTech.Cells(1, vStartColIndex).Value = "Îñíîâíîé ïðèõîä"
        sheetTech.Cells(2, vStartColIndex).Value = cell.Value
        vStartColIndex = vStartColIndex + 1
    Next
    
    
    ' Lecturer sums
    For Each cell In sheetCosts.Range("F49", "F58").Cells
        sheetTech.Cells(1, vStartColIndex).Value = "Ëåêòîðñêèå|Ñóììà"
        sheetTech.Cells(2, vStartColIndex).Value = cell.Value
        vStartColIndex = vStartColIndex + 1
    Next
    

    ' Fees
    For Each cell In sheetCosts.Range("H49", "H58").Cells
        sheetTech.Cells(1, vStartColIndex).Value = "Êîìèññèÿ"
        sheetTech.Cells(2, vStartColIndex).Value = cell.Value
        vStartColIndex = vStartColIndex + 1
    Next
    
    
    ' Legal entities
    For Each cell In sheetCosts.Range("J49", "J58").Cells
        sheetTech.Cells(1, vStartColIndex).Value = "Þðëèöî"
        sheetTech.Cells(2, vStartColIndex).Value = cell.Value
        vStartColIndex = vStartColIndex + 1
    Next
    
    MsgBox vStartColIndex
End Sub

Sub copyFromTechToTable()
    MsgBox "Start copyFromTechToTable"
    Dim sheetTech As Worksheet
    Dim sheetCosts As Worksheet
    Dim sheetTable As Worksheet
    
    Set sheetTech = getOrCreateSheet("òåõí")
    Set sheetCosts = getOrCreateSheet("ñìåòà")
    Set sheetTable = getOrCreateSheet("òàáëèöà")
    
    ' Headers
    copyValue sheetTech.Range("B1", "G1"), sheetTable.Range("B1", "G1")
    
    ' Values
    copyValue sheetTech.Range("B2", "G2"), sheetTable.Range("B2", "G2")
    
    Dim sumCash As Double
    Dim sumNonCash As Double
    
    sumCash = 0#
    sumNonCash = 0#
    
    For Each cell In sheetTech.Range("H2", "AL2").Cells
        If cell.Value <> "" Then
            vals = Split(cell.Value, "::")
            
            If vals(0) = "ÁÍ" Then
                sumNonCash = sumNonCash + CDbl(vals(1))
            ElseIf vals(0) = "Í" Then
                sumCash = sumCash + CDbl(vals(1))
            Else
                Err.Raise -20001, "Êîïèðâàíèå Òåõí â Òàáëèöà", "Çíà÷åíèå îòëè÷íîå îò ÁÍ è Í äëÿ ñóìì ñìåòû"
            End If
            
        End If
    Next
    
    sheetTable.Range("H1").Value = "Ðàñõîäû á/í"
    sheetTable.Range("H2").Value = sumNonCash
    
    sheetTable.Range("I1").Value = "Ðàñõîäû í"
    sheetTable.Range("I2").Value = sumCash
    
    
    
    Dim companies As String
    
    companies = ""
    
    For Each cell In sheetTech.Range("AM2", "AV2").Cells
        If cell.Value <> "" Then
            If companies = "" Then
                companies = cell.Value
            Else
                companies = companies & ", " & cell.Value
            End If
            
        End If
    Next
    
    sheetTable.Range("J1").Value = "Íàçâàíèå êîìïàíèè"
    sheetTable.Range("J2").Value = companies
    
        
    Dim mainIncome As Double
    mainIncome = 0#
    
    For Each cell In sheetTech.Range("AW2", "BF2").Cells
        If cell.Value <> "" Then
            mainIncome = mainIncome + CDbl(cell.Value)
            
        End If
    Next
    
    sheetTable.Range("K1").Value = "Îñíîâíîé ïðèõîä"
    sheetTable.Range("K2").Value = mainIncome
    
    
    Dim sumLecturers As Double
    sumLecturers = 0#
    
    For Each cell In sheetTech.Range("BG2", "BP2").Cells
        If cell.Value <> "" Then
            sumLecturers = sumLecturers + CDbl(cell.Value)
            
        End If
    Next
    
    sheetTable.Range("L1").Value = "Ëåêòîðñêèå"
    sheetTable.Range("L2").Value = sumLecturers
    
    
    Dim fees As Double
    fees = 0#
    
    For Each cell In sheetTech.Range("BQ2", "BZ2").Cells
        If cell.Value <> "" Then
            fees = fees + CDbl(cell.Value)
            
        End If
    Next
    
    sheetTable.Range("M1").Value = "Êîìèññèÿ"
    sheetTable.Range("M2").Value = fees
    
    
    Dim legalEntities As String
    
    legalEntities = ""
    
    For Each cell In sheetTech.Range("CA2", "CJ2").Cells
        If cell.Value <> "" Then
            If legalEntities = "" Then
                legalEntities = cell.Value
            Else
                legalEntities = legalEntities & ", " & cell.Value
            End If
            
        End If
    Next
    
    sheetTable.Range("N1").Value = "Þðëèöà"
    sheetTable.Range("N2").Value = legalEntities
    
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

Function getNextId()
    Dim sheetTech As Worksheet
    Set sheetTech = getOrCreateSheet("òåõí")
    
    sheetTech.Range("B2").Value = sheetTech.Range("B2").Value + 1
    getNextId = sheetTech.Range("B1").Value & "_" & sheetTech.Range("B2").Value
    
    MsgBox getNextId
End Function

Function getLastId()
    Dim sheetTech As Worksheet
    Set sheetTech = getOrCreateSheet("òåõí")
    
    getLastId = sheetTech.Range("B1").Value & "_" & sheetTech.Range("B2").Value
    
    MsgBox getLastId
End Function

