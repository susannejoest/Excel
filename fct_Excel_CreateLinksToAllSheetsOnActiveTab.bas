Sub CreateLinksToAllSheetsOnActiveTab()

    Dim xlWs As Worksheet
    Dim strWsName As String
    
    'This line declares xlWs as a worksheet object.
    
    Dim cell As Range
    
    'Cell is declared a range object. A range object can contain a single cell or multiple cells.
    
    For Each xlWs In ActiveWorkbook.Worksheets
    
    'Each worksheet in active workbook is stored in sh, one by one.
    
    If ActiveSheet.Name <> xlWs.Name Then
    
    'This If ... then line avoids linking to active worksheet.
    strWsName = "'" & xlWs.Name & "'"
    
    If ActiveCell.Value > "" Then
        MsgBox "Next Cell down not empty!"
        Exit Sub
    End If
    
    ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
    strWsName & "!A1", TextToDisplay:=xlWs.Name
     
    'Create a hyperlink to current worksheet sh in active cell.
    
    ActiveCell.Offset(1, 0).Select
    
    'Select next cell below active cell.
    
    End If
    
    Next xlWs
    
    'Go back to the "For each" statement and store next worksheet in sh worksheet object.

End Sub

