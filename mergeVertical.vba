Public Sub mergeVertical(control As IRibbonControl)

Dim current As Worksheet

    Dim data As range
    Dim insertData As range
    Dim tempStr As String
    Dim longStr As String
    Dim lastRow As Integer
    Dim lastColumn As Integer
    Dim answer As Integer
        answer = MsgBox("Are you sure you want to merge?", vbYesNo + vbQuestion, "Empty Sheet")
    If answer = vbYes Then
    Set data = Selection
    
    'MsgBox data.Address Caly adres
    'MsgBox data.Column Pierwsza kolumna
    'MsgBox data.Row Pierwszy rzad W kazdym row bedzie merge columns
    
    Set current = Application.ActiveSheet
    
   lastRow = (data.Row + data.rows.Count - 1)
   lastColumn = (data.Column + data.columns.Count - 1)
        
    For i = 0 To data.rows.Count - 1
    For j = 0 To data.columns.Count - 1
    tempStr = range(current.cells(data.Row + i, data.Column + j), current.cells(data.Row + i, data.Column + j)).Text
    'MsgBox "For i: " & i & " and j: " & j & " :-: " & tempStr
    longStr = longStr & " " & tempStr
    range(current.cells(data.Row + i, data.Column + j), current.cells(data.Row + i, data.Column + j)).Value = ""
    Next
    'MsgBox "NEXT: " & longStr
    Set insertData = current.range(current.cells(data.Row + i, data.Column), current.cells(data.Row + i, data.Column))
    insertData.Value = longStr
    
    longStr = ""
    Next
    Else: End If

End Sub
