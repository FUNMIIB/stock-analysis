Sub DQAnalysis()
        Worksheets("DQ Analysis").Activate
        
        Range("A1").Value = "DAQ0 (Ticker: DQ)"
        
        'Create a header row
        Cells(3, 1).Value = "Year"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"
        
        Worksheets("2018").Activate
        
        'set initial volume to zero
    totalVolume = 0
    
    Dim startingPrice As Double
    Dim endingPrice As Double
    'Establish the number of rows to loop over
rowStart = 2
RowEnd = Cells(Rows.Count, "A").End(xlUp).Row

        'loop over all the rows
    For i = rowStart To RowEnd
            
            If Cells(i, 1).Value = "DQ" Then
            
            'increase totalVolume by the value in the current row
            totalVolume = totalVolume + Cells(i, 8).Value
            
    End If
    
        If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
                
                startingPrice = Cells(i, 6).Value
    End If
        If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
        
                endingPrice = Cells(i, 6).Value
    End If
    
    Next i
    
    Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = (endingPrice / startingPrice) - 1
    
End Sub
                
    
        



