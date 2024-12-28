Sub Module2Challenge()
    ' Worksheet variables
    Dim ws As Worksheet
    Dim lastRow As Long

    ' Tracking variables
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double

    ' Summary table variables
    Dim summaryRow As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim increaseTickerName As String
    Dim decreaseTickerName As String
    Dim volumeTickerName As String

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Variables
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        summaryRow = 2
        totalVolume = 0
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0

        ' Headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        ' First ticker and open price
        ticker = ws.Cells(2, 1).Value
        openPrice = ws.Cells(2, 3).Value

        ' Data loop
        Dim i As Long
        For i = 2 To lastRow

            ' Accumulate total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value

            ' Check if the current row is the last for the ticker
            If ws.Cells(i + 1, 1).Value <> ticker Or i = lastRow Then
                ' Closing price
                closePrice = ws.Cells(i, 6).Value

                ' Calculate change
                yearlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = yearlyChange / openPrice
                Else
                    percentChange = 0
                End If

                ' Summary table
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = yearlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalVolume

                ' Cell format
                ws.Cells(summaryRow, 11).NumberFormat = "0.00%"

                ' Conditional formatting for yearly change
                If yearlyChange > 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
                ElseIf yearlyChange < 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
                Else
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 255, 255) ' White
                End If

                ' Check for greatest values
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    increaseTickerName = ticker
                End If

                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    decreaseTickerName = ticker
                End If

                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    volumeTickerName = ticker
                End If

                ' Move to the next summary row
                summaryRow = summaryRow + 1

                ' Reset variables for the next ticker
                ticker = ws.Cells(i + 1, 1).Value
                openPrice = ws.Cells(i + 1, 3).Value
                totalVolume = 0
            End If
        Next i

        ' Create Greatest Values Summary
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        ' Fill in greatest values
        ws.Cells(2, 16).Value = increaseTickerName
        ws.Cells(2, 17).Value = greatestIncrease
        ws.Cells(3, 16).Value = decreaseTickerName
        ws.Cells(3, 17).Value = greatestDecrease
        ws.Cells(4, 16).Value = volumeTickerName
        ws.Cells(4, 17).Value = greatestVolume
    Next ws
End Sub
