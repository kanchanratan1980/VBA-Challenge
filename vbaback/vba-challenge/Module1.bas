Attribute VB_Name = "Module1"
Sub StockAnalysis()

    Dim ws As Worksheet
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim startRow As Long
    Dim lastRow As Long
    Dim summaryRow As Long
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Initialize variables
        summaryRow = 2
        maxIncrease = 0
        maxDecrease = 0
        maxVolume = 0

        ' Find the last row of data
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Add headers for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"

        ' Loop through rows to analyze data
        For i = 2 To lastRow
            ' Check if a new ticker starts
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                startRow = i
                totalVolume = 0
            End If

            ' Accumulate total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value

            ' Check if the current ticker ends
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(startRow, 3).Value
                closePrice = ws.Cells(i, 6).Value

                ' Calculate quarterly change and percent change
                quarterlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = (quarterlyChange / openPrice)
                Else
                    percentChange = 0
                End If

                ' Output to summary table
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = quarterlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalVolume

                ' Apply conditional formatting
                If quarterlyChange > 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0)
                Else
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0)
                End If
                If percentChange > 0 Then
                    ws.Cells(summaryRow, 11).Interior.Color = RGB(0, 255, 0)
                Else
                    ws.Cells(summaryRow, 11).Interior.Color = RGB(255, 0, 0)
                End If

                ' Update greatest values
                If percentChange > maxIncrease Then
                    maxIncrease = percentChange
                    maxIncreaseTicker = ticker
                End If
                If percentChange < maxDecrease Then
                    maxDecrease = percentChange
                    maxDecreaseTicker = ticker
                End If
                If totalVolume > maxVolume Then
                    maxVolume = totalVolume
                    maxVolumeTicker = ticker
                End If

                ' Move to the next summary row
                summaryRow = summaryRow + 1
            End If
        Next i

        ' Format Percent Change column
        ws.Columns(11).NumberFormat = "0.00%"

        ' Output greatest values
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(2, 16).Value = maxIncreaseTicker
        ws.Cells(3, 16).Value = maxDecreaseTicker
        ws.Cells(4, 16).Value = maxVolumeTicker
        ws.Cells(2, 17).Value = maxIncrease
        ws.Cells(3, 17).Value = maxDecrease
        ws.Cells(4, 17).Value = maxVolume

    Next ws

End Sub
