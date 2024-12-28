Sub Module 2 Challenge()
    ' Worksheet variables
    Dim as Worksheet
    Dim lastRow as Long

    ' Tracking variables
    Dim ticker as String
    Dim openPrice as Double
    Dim closePrice as Double
    Dim yearlyChange as Double
    Dim percentChange as Double
    Dim totalVolume as Double

    ' Summary table variables
    Dim summaryRow as Long
    Dim greatetIncrease as Double
    Dim greatestDecrease as Double
    Dim greatestVolume as Double
    Dim invreaseTickerName as String
    Dim decreaseTickerName as String
    Dim volumeTickerName as String

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
    ' Variables
    lastRow = ws.Cells(Rows.count, 1).End(xlUp).row 
    summaryRow = 2
    totalVolume = 0
    greatetIncrease = 0
    greatestDecrease = 0
    greatestVolume = =0

    ' Headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'First ticker and open price
    ticker = ws.Cells(2, 1).Value
    openPrice = ws.Cells(2, 3).Value

    'Data loop
    Dim i as Long
    For i = 2 to lastRow

        totalVolume = totalVolume + ws.Cells(i, 7).Value

        ' last row of current ticker
        If ws.Cells(i + 1, 1).Value <> ticker Then

            'Closing price
            closePrice = ws.cells(i, 6).Value

            'Calculate change
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
              ws.cells(summaryRow, 11).NumberFormat = "0.00%" 

            ' Conditional formatting for yearly change
            If yearlyChange > 0 Then
            ' Positive change - Green
                ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
            ElseIf yearlyChange < 0 Then
            ' Negative change - Red
                ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
            Else
            ' No change - White (neutral)
                ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 255, 255) ' White
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

        





            







