Sub hello()
    MsgBox "Start"
    Dim sheetTech As Worksheet
    Dim sheetÑosts As Worksheet
    Dim sheetTable As Worksheet
    
    Set sheetTech = getOrCreateSheet("òåõí")
    Set sheetÑosts = getOrCreateSheet("ñìåòà")
    Set sheetTable = getOrCreateSheet("òàáëèöà")
    
    Dim rngFrom As Range
    Dim rngTo As Range
    
    'Set rngFrom = sheetÑosts.Range("A3", "A8")
    'Set rngTo = sheetTech.Range("B1", "I1")
    
    'rngTo.Value2 = rngFrom.Value2
    
    copyValue sheetÑosts.Range("A3", "A8"), sheetTech.Range("B1", "I1")
    
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


