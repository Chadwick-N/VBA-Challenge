Sub Ticker()

    Dim Ticker As String
    Dim Year_Change As Double
    Dim Percent_Change As Double
    Dim Total_Vol As LongLong
    Dim Summary_Table_Row As Long
    Dim Lastrow As Long

    Summary_Table_Row = 2
    start_row = 2
    
        For Each ws In Worksheets
        
        Range("J1") = "Ticker"
        Range("K1") = "Yearly Change"
        Range("L1") = "Percentage Change"
        Range("M1") = "Total Volume"
        
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            For i = 2 To Lastrow
                
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    start_value = ws.Cells(start_row, 3).Value
                    end_value = ws.Cells(i, 6).Value
                    Year_Change = end_value - start_value
                    Ticker = Cells(i, 1).Value
                    Total_Vol = Total_Vol + Cells(i, 7).Value
                    Percent_Change = Round((Year_Change / start_value) * 100, 2)
                    
                    ws.Range("J" & Summary_Table_Row).Value = Ticker
                    ws.Range("K" & Summary_Table_Row).Value = Year_Change
                    ws.Range("L" & Summary_Table_Row).Value = Percent_Change
                    ws.Range("M" & Summary_Table_Row).Value = Total_Vol
                    
                    Summary_Table_Row = Summary_Table_Row + 1
                    Total_Vol = 0
                    
                Else:
                Total_Vol = Total_Vol + ws.Cells(i, 7).Value
                        
                End If
                
            Next i
        
        For j = 2 To Lastrow
        Year_Change = Cells(j, 11).Value
        
            If Year_Change < 0 Then
                Cells(j, 11).Interior.ColorIndex = 3
                        
                Else
                    Cells(j, 11).Interior.ColorIndex = 4
                    
            End If
            
        Next j
        
    Next ws
    
    
End Sub


