Sub AlphabeticalTest()
    For Each Ws In Worksheets
    
    
        Ws.Range("J1").Value = "Ticker"
        Ws.Range("K1").Value = "Totale Stock Volume"
        
        Dim Ticker As String
        Dim TotalVol As Double
        TotalVol = 0
        
        Dim SummaryTable As Integer
        SummaryTable = 2
        
        LastRow = Ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To LastRow
            
            If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
                Ticker = Ws.Cells(i, 1).Value
                TotalVol = TotalVol + Ws.Cells(i, 7).Value
                
                Ws.Range("J" & SummaryTable).Value = Ticker
                Ws.Range("K" & SummaryTable).Value = TotalVol
                
                SummaryTable = SummaryTable + 1
                TotalVol = 0
                Else
            
                TotalVol = TotalVol + Ws.Cells(i, 7).Value
            
            End If
        Next i
    Next Ws

End Sub