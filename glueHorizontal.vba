
Public Sub glueHzn(control As IRibbonControl)

    Dim data As range
    Dim insertData As range
    Dim tempStr As String
    Dim longStr As String
    Dim lastRow As Integer
    Dim lastColumn As Integer
    Dim answer As Integer
    
    Set data = Selection
    
    Set current = Application.ActiveSheet
    
   lastRow = (data.Row + data.rows.Count - 1)
   lastColumn = (data.Column + data.columns.Count - 1)

    For i = 0 To data.rows.Count - 2
    For j = 0 To data.columns.Count - 1
    tempStr = range(current.cells(data.Row + i, data.Column + j), current.cells(data.Row + i, data.Column + j)).Text
    If tempStr = "" Then
    range(current.cells(data.Row + i, data.Column + j), current.cells(data.Row + i, data.Column + j)).Value = range(current.cells(data.Row + i + 1, data.Column + j), current.cells(data.Row + i + 1, data.Column + j)).Text
    range(current.cells(data.Row + i + 1, data.Column + j), current.cells(data.Row + i + 1, data.Column + j)).Value = ""
    End If
    Next
    Next
    
    
    
End Sub
