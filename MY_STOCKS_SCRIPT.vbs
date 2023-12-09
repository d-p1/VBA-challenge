Sub multi_stocks()

Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"

Dim Ticker As String

Dim open_price As Double

Dim close_price As Double

Dim Total_Stock_Volume As LongLong

Dim Ticker_Row As Integer

open_price = ws.Cells(2, 3)
close_price = 0
Total_Stock_Volume = 0
Ticker_Row = 2

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow
    If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
        Ticker = ws.Cells(i, 1)
        close_price = ws.Cells(i, 6)
        Percent_Change = (close_price - open_price) / open_price
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7)
        ws.Range("I" & Ticker_Row) = Ticker
        ws.Range("J" & Ticker_Row) = close_price - open_price
        ws.Range("J" & Ticker_Row).NumberFormat = "$0.00; -$0.00"
        ws.Range("K" & Ticker_Row) = Percent_Change
        ws.Range("K" & Ticker_Row).NumberFormat = "0.00%"
        ws.Range("L" & Ticker_Row) = Total_Stock_Volume
        
            If ws.Range("J" & Ticker_Row) > 0 Then
                ws.Range("J" & Ticker_Row).Interior.ColorIndex = 4
            ElseIf ws.Range("J" & Ticker_Row) <= 0 Then
                ws.Range("J" & Ticker_Row).Interior.ColorIndex = 3
            End If
            
            If ws.Range("K" & Ticker_Row) > 0 Then
                ws.Range("K" & Ticker_Row).Interior.ColorIndex = 4
            ElseIf ws.Range("K" & Ticker_Row) <= 0 Then
                ws.Range("K" & Ticker_Row).Interior.ColorIndex = 3
            End If
        open_price = ws.Cells((i + 1), 3)
        Total_Stock_Volume = 0
        close_price = 0
        Ticker_Row = Ticker_Row + 1

    Else
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7)
    End If
Next i

Dim greatest_increase As Double
Dim greatest_decrease As Double
Dim greatest_volume As LongLong


ws.Range("Q2") = WorksheetFunction.Max(ws.Range("K2" & ":" & "K" & lastrow))
ws.Range("Q3") = WorksheetFunction.Min(ws.Range("K2" & ":" & "K" & lastrow))
ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2" & ":" & "L" & lastrow))

ws.Range("Q1") = "Value"
ws.Range("P1") = "Ticker"
ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total Volume"

greatest_increase = ws.Range("Q2")
greatest_decrease = ws.Range("Q3")
greatest_volume = ws.Range("Q4")

ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").NumberFormat = "0.00%"


    For j = 2 To lastrow
        If ws.Cells(j, 11) = greatest_increase Then
            ws.Range("P2") = ws.Cells(j, 9)
        
        ElseIf ws.Cells(j, 11) = greatest_decrease Then
            ws.Range("P3") = ws.Cells(j, 9)
    
        ElseIf ws.Cells(j, 12) = greatest_volume Then
            ws.Range("P4") = ws.Cells(j, 9)
        End If
    Next j
Next ws

End Sub
