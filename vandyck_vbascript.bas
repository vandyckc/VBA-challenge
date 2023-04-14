Attribute VB_Name = "Module1"
Sub ABTesting():

'assign variables
Dim ticker As String
Dim yearchange As Double
Dim openprice As Double
Dim closeprice As Double
Dim tot_volume As Double
Dim perc_change As Double
Dim max_ticker As String
Dim max_volume As Double
    Total = 0
    max_volume = 0

For Each ws In Worksheets

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
summary_row = 2

'create column labels
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"
    ws.Cells(2, 15) = "Greatest % Increase"
    ws.Cells(3, 15) = "Greatest % Decrease"
    ws.Cells(4, 15) = "Greatest Total Volume"
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"
    
'loop through for ticker, yearly change, percent change, total volume
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
                ticker = ws.Cells(i, 1)
                tot_volume = tot_volume + ws.Cells(i, 7)
                    ws.Cells(summary_row, 9) = ticker
                    ws.Cells(summary_row, 12) = tot_volume
                tot_volume = 0
                ticker = ""
                closeprice = ws.Cells(i, 6)
                yearchange = closeprice - openprice
                            ws.Cells(summary_row, 10) = yearchange
                perc_change = (yearchange / openprice)
                            ws.Cells(summary_row, 11) = perc_change
                summary_row = summary_row + 1
            Else
                If ws.Cells(i - 1, 1) <> ws.Cells(i, 1) Then
                openprice = ws.Cells(i, 3)
                End If
                tot_volume = tot_volume + ws.Cells(i, 7)
            End If
        Next i
        LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row

'loop through to analyze max volume, greatest increase, greatest decrease
        For i = 2 To LastRow2
            If ws.Cells(i, 12) > max_volume Then
                max_volume = ws.Cells(i, 12)
                max_ticker = ws.Cells(i, 9)
            End If
            If ws.Cells(i, 11) > greatest_increase Then
                greatest_increase = ws.Cells(i, 11)
                ticker_greatest_increase = ws.Cells(i, 9)
            End If
            If ws.Cells(i, 11) < greatest_decrease Then
                greatest_decrease = ws.Cells(i, 11)
                ticker_greatest_increase = ws.Cells(i, 9)
            End If
        Next i

'assign analysis column labels
    ws.Range("Q4") = max_volume
    ws.Range("P4") = max_ticker
    ws.Range("Q2") = greatest_increase
    ws.Range("P2") = ticker_greatest_increase
    ws.Range("Q3") = greatest_decrease
    ws.Range("P3") = ticker_greatest_decrease

'variable reset for next ws
    Ticker_Row = 2
    max_volume = 0
    max_ticker = ""

'cell formatting
    ws.Range("K:K").NumberFormat = "0.00%"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    ws.Range("I:L").Columns.AutoFit
    ws.Range("O:Q").Columns.AutoFit
        For i = 2 To LastRow2
            If ws.Cells(i, 10) < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
            Else
            ws.Cells(i, 10).Interior.ColorIndex = 4
            End If
        Next i
Next ws
End Sub
