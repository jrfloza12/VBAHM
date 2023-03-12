Sub Stock_Test2.2()

Dim ws As Worksheet


For Each ws In Worksheets
ws.Activate

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("i1").Value = "Ticker"
    Range("j1").Value = "Yearly Change"
    Range("k1").Value = "Percent Change"
    Range("l1").Value = "Total Stock Volume"
    VOL = 0
    TR = 2
    
    
    
    For R = 2 To lastRow
        
        VOL = VOL + Cells(R, 7).Value
        
       If ws.Cells(R, 1).Value <> ws.Cells(R - 1, 1).Value Then
            OP = ws.Cells(R, 3).Value
            
        End If
       
       
       
        If ws.Cells(R, 1) <> ws.Cells(R + 1, 1).Value Then
            TK = ws.Cells(R, 1).Value
            ws.Cells(TR, 9).Value = TK
            
            
            CP = Cells(R, 6).Value
            
            YearlyChange = CP - OP
            PercentChange = YearlyChange / OP
            
            TotalStockVolume = VOL
            
            Range("k2:k" & lastRow).NumberFormat = "0.00%"
            Cells(TR, 10).Value = YearlyChange
            Cells(TR, 11).Value = PercentChange
            Cells(TR, 12).Value = TotalStockVolume
            
        
            
            TR = TR + 1
            VOL = 0
        
        End If
        
    
    Next R
    
    
    
    
    
    
    
    
Next ws

End Sub

