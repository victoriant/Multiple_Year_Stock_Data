Sub Bonus()

'Bonus:

Dim mxdiff = WorksheetFuntion.Max(Range("K2:K10000"))
Dim mndiff = WorksheetFuntion.Min(Range("K2:K10000"))
Dim vldiff = WorksheetFuntion.Max(Range("K2:K10000"))

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"


    For j = 2 To LastRow
    
        If mxdiff = Cells(i, 12).Value Then
        Range("P2") = Cells(i, 9).Value
        Else if mndiff = Cells(i,12) Then
        Range("P3") = Cells(i, 9).Value
        Else vldiff = Cells(1,13).Value Then
        Range("P4").Value = Cells(i, 9).Value
        
        End If
        
    Next j
    
Next ws

End Sub
