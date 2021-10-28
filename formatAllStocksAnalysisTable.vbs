Sub formatAllStocksAnalysisTable()
 'Formatting
        Worksheets("All Stocks Analysis").Activate
    'to change font to bold
        Range("A3:C3").Font.Bold = True
    'to create borders on edge bottom
        Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinous
    'to change font color
        Range("A3:C3").Font.Color = vbRed
    'to change background color
        Range("A3:C3").Interior.ColorIndex = 4
        Range("B4:B15").NumberFormat = "#,##0"
        Range("C4:C15").NumberFormat = "0.0%"
        Range("C4:C15").NumberFormat = "0.00%"
        Range("B4:B15").NumberFormat = "$#,##0.00"
        Columns("B").AutoFit
        
 dataRowStart = 4
 dataRowEnd = 15
 For i = dataRowStart To dataRowEnd
 
If Cells(i, 3) > 0 Then
    'Color the cell green
    Cells(i, 3).Interior.Color = vbGreen
ElseIf Cells(i, 3) < 0 Then

    'Color the cell red
    Cells(i, 3).Interior.Color = vbRed
    
Else
    'Clear the cell color
    Cells(i, 3).Interior.Color = xlNone
    
End If
Next i
        

        
        
        
        
        
        
        
End Sub
