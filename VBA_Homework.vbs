Sub alphabetic_testing_hw()

Dim ws As Worksheet
    For Each ws In Worksheets
    ws.Activate

        Dim total_v As Long
        Dim ticker_n As String
        
        Cells(1, "i").Value = "Ticker Symbol"
        Cells(1, "j").Value = "Total Stock Volume"
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
        v_count = 2
    
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker_n = ws.Cells(i, "a").Value
                total_v = ws.Cells(i, "g").Value
                ws.Cells(v_count, "i").Value = ticker_n
                ws.Cells(v_count, "j").Value = total_v
                v_count = v_count + 1
                total_v = 0
            End If
        Next i
    Next ws

End Sub



