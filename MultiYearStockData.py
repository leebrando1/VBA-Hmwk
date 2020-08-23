Sub stockdata():

    Dim ticker As String
    Dim openprice As Double
    Dim closeprice As Double
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim totalvolume As Double

    Dim greatest_increase_ticker As String
    Dim greatest_decrease_ticker As String
    Dim greatest_volume_ticker As String

    For Each ws In Worksheets

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

    Dim openprice_row As Long
    openprice_row = 2
    totalvolume = 0

    Dim summary_table_row As Integer
    summary_table_row = 2

    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastrow

        'the last cell of each unique stock ticker
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            'ticker and total volume values
            ticker = ws.Cells(i, 1).Value
            totalvolume = totalvolume + ws.Cells(i, 7).Value

            'print ticker
            ws.Range("I" & summary_table_row).Value = ticker

            'print total volume
            ws.Range("L" & summary_table_row).Value = totalvolume

            'yearly change
            openprice = ws.Range("C" & openprice_row).Value
            closeprice = ws.Range("F" & i).Value
            yearlychange = closeprice - openprice

                'percent change
                If openprice = 0 Then
                    percentchange = 0
                Else
                    percentchange = yearlychange / openprice
                End If

            'print yearly change and percent change
            ws.Range("J" & summary_table_row).Value = yearlychange
            ws.Range("k" & summary_table_row).Value = percentchange
            ws.Range("k" & summary_table_row).NumberFormat = "0.00%"
            ws.Range("J" & summary_table_row).NumberFormat = "0.00"

                'Red if neg and Green if pos
                If ws.Range("J" & summary_table_row).Value > 0 Then
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
                End If

            summary_table_row = summary_table_row + 1
            openprice_row = i + 1
            totalvolume = 0
        Else
            totalvolume = totalvolume + ws.Range("G" & i).Value

    End If

Next i


'lastrow for consolidated tickers
lastrow_greatest = ws.Cells(Rows.Count, "I").End(xlUp).Row

greatest_increase = 0
greatest_decrease = 0
greatest_volume = 0

For j = 2 To lastrow_greatest
    If Range("K" & j + 1).Value > greatest_increase Then
        greatest_increase = ws.Range("K" & j + 1).Value
        greatest_increase_ticker = ws.Range("I" & j + 1).Value
    Else
        greatest_increase = greatest_increase
    End If

    If ws.Range("K" & j + 1) < greatest_decrease Then
        greatest_decrease = ws.Range("K" & j + 1).Value
        greatest_decrease_ticker = ws.Range("I" & j + 1).Value
    Else
        greatest_decrease = greatest_decrease
    End If

    If ws.Range("L" & j + 1).Value > greatest_volume Then
        greatest_volume = ws.Range("L" & j + 1).Value
        greatest_volume_ticker = ws.Range("I" & j + 1).Value
    Else
        greatest_volume = greatest_volume
    End If

Next j

    'print greatest values
    ws.Range("Q2").Value = greatest_increase
    ws.Range("Q3").Value = greatest_decrease
    ws.Range("Q4").Value = greatest_volume
    ws.Range("P2").Value = greatest_increase_ticker
    ws.Range("P3").Value = greatest_decrease_ticker
    ws.Range("P4").Value = greatest_volume_ticker
    ws.Range("Q2:Q3").NumberFormat = "0.00%"

Next ws

End Sub
