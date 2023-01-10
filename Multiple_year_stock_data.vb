
Dim ws As Worksheet
    
For Each ws In Worksheets
        
Dim tickername As String
'Ticker = " "
Dim tickervolume As Double
tickervolume = 0
Dim summary_ticker_row As Integer
summary_ticker_row = 2
    
Dim open_price As Double
open_price = ws.Cells(2, 3).Value
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double
    
Dim Max As Double
Dim Min As Double
Dim j As Long
Dim i As Long
'Dim lastrow As Long
    
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
        
    
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        tickername = ws.Cells(i, 1).Value
            
        tickervolume = tickervolume + ws.Cells(i, 7).Value
            
        ws.Range("I" & summary_ticker_row).Value = tickername
        ws.Range("L" & summary_ticker_row).Value = tickervolume
        
        close_price = ws.Cells(i, 6).Value
            
        yearly_change = (close_price - open_price)
            
        ws.Range("J" & summary_ticker_row).Value = yearly_change
            
             If (open_price = 0) Then
                percent_change = 0
                
            Else
                percent_change = yearly_change / open_price
                
            End If
            
        ws.Range("K" & summary_ticker_row).Value = percent_change
        ws.Range("K" & summary_ticker_row).NumberFormat = "0.00%"
            
        summary_ticker_row = summary_ticker_row + 1
            
        tickervolume = 0
            
        open_price = ws.Cells(i + 1, 3)
            
        Else
            
        tickervolume = tickervolume + ws.Cells(i, 7).Value
            
        End If
            
            If ws.Cells(i, 10).Value > 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 10
               
            Else
                    ws.Cells(i, 10).Interior.ColorIndex = 3
               
            
            End If
            
    Next i

Next ws
    

End Sub