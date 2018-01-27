Sub Final():
For Each WS In Worksheets
WS.Select
paste_stock_volume
yearly_change
Next WS
End Sub


Sub paste_stock_volume():
Dim stock_name As String
Dim Summary_Table_Row As Double
Dim stock_total As Double
For Each WS In Worksheets
    WS.Select
    Cells(1, 9) = "Ticker"
    Cells(1, 12) = "Total Volume"
    Summary_Table_Row = 2
    last_row = WS.Cells(Rows.Count, "A").End(xlUp).Row
    For i = 2 To last_row
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            stock_name = Cells(i, 1).Value
            stock_total = stock_total + Cells(i, 7)
            Range("I" & Summary_Table_Row).Value = stock_name
            Range("L" & Summary_Table_Row).Value = stock_total
            Summary_Table_Row = Summary_Table_Row + 1
            stock_total = 0
        Else
            stock_total = stock_total + Cells(i, 7).Value
        End If
    Next i
Next WS
End Sub

Sub yearly_change():
Dim stock_first As Double
Dim Summary_Table_Row As Double
Dim stock_last As Double
Dim stock_count As Double
Dim stock_change As Double
For Each WS In Worksheets
    WS.Select
    Cells(1, 10) = "Yr Change"
    Cells(1, 11) = "% Change"
    Summary_Table_Row = 2
    last_row = WS.Cells(Rows.Count, "A").End(xlUp).Row
    For i = 2 To last_row
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            stock_count = (stock_count + Cells(i, 1).Count) - 1
            stock_first = Cells(i - stock_count, 3).Value
            stock_last = Cells(i, 6).Value
            stock_change = stock_last - stock_first
            Range("J" & Summary_Table_Row).Value = stock_change
            If stock_first <> 0 Then
                Range("K" & Summary_Table_Row).Value = stock_change / stock_first
            Else
                Range("K" & Summary_Table_Row).Value = 0
            End If
            Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            If Range("J" & Summary_Table_Row).Value < 0 Then
                Range("J" & Summary_Table_Row).Interior.Color = vbRed
            Else
                Range("J" & Summary_Table_Row).Interior.Color = vbGreen
            End If
            Summary_Table_Row = Summary_Table_Row + 1
            stock_count = 0
        Else
        stock_count = stock_count + Cells(i, 3).Count
       End If
    Next i
Next WS
End Sub