Option Explicit

Public Sub UniqueTicker()
'variable assignment
Dim tickerName As String
Dim totalTickerVolumePerTicker As LongLong
Dim row As LongLong
Dim uniqueTickerRow As Integer
Dim isFirstOpenPrice As Boolean
Dim differenceInPrice As Double
Dim firstOpenPrice As Double
Dim lastClosePrice As Double
Dim percentageChange As Double
Dim greatestPercentageIncrease As Double
Dim greatestPercentageTickerName As String
Dim greatestDecreasePercentageTickerName As String
Dim greatestPercentageDecrease As Double
Dim greatestTotalTickerVolumeName As String
Dim greatestTotalTickerVolume As LongLong
Dim uniqueTickertRow2 As Integer
Dim ws As Worksheet

'here the for loop will iterate through all the worksheets and execute the following logic on each of them
For Each ws In Worksheets
'reset the variable for each worksheet
    totalTickerVolumePerTicker = 0
    row = 2
    isFirstOpenPrice = True
    uniqueTickerRow = 2
    uniqueTickertRow2 = 2
   'name the columns
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change($)"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 16).Value = "Ticker"
    ws.Cells(2, 17).Value = "Value"
   'below, the do until loop will perform the code repetatively unless it sees an empty cell !
   'Once it sees the empty cell, it will go to the another sheet and start to perform the code.
    Do Until IsEmpty(ws.Cells(row, 3))
          ' if the logic below finds the unique ticker, it calculates sum of volume for each ticker symbol
          ' else it adds up with the same ticker volume
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
                tickerName = ws.Cells(row, 1).Value
                totalTickerVolumePerTicker = totalTickerVolumePerTicker + ws.Cells(row, 7).Value
                ws.Cells(uniqueTickerRow, 9).Value = tickerName
                ws.Cells(uniqueTickerRow, 12).Value = totalTickerVolumePerTicker
                ' checks if the row is first for the Ticekr
                ' and if "Yes", store the opening price
                If isFirstOpenPrice = True Then
                    firstOpenPrice = ws.Cells(row, 3).Value
                End If
                
                lastClosePrice = ws.Cells(row, 6).Value
                differenceInPrice = lastClosePrice - firstOpenPrice
                ws.Cells(uniqueTickerRow, 10).Value = differenceInPrice
               'the logic below compares the "Yearly Change" price with 0 to find out the negative and positve price and highlights in red and green respectively.
                If differenceInPrice < 0 Then
                    ws.Cells(uniqueTickerRow, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(uniqueTickerRow, 10).Interior.ColorIndex = 4
                End If
                'uses the initial opening price and the end closing price in each year, total them and find the Yearly change in price
                percentageChange = differenceInPrice / firstOpenPrice
                ws.Cells(uniqueTickerRow, 11).Value = percentageChange
                ws.Cells(uniqueTickerRow, 11).NumberFormat = "0.00%"
                uniqueTickerRow = uniqueTickerRow + 1
                totalTickerVolumePerTicker = 0
                row = row + 1
                isFirstOpenPrice = True
           Else
                totalTickerVolumePerTicker = totalTickerVolumePerTicker + ws.Cells(row, 7).Value
                If isFirstOpenPrice = True Then
                    firstOpenPrice = ws.Cells(row, 3).Value
                    isFirstOpenPrice = False
                End If
                row = row + 1
           End If
    Loop
    'value assigned to the variable
    greatestPercentageIncrease = ws.Cells(uniqueTickertRow2, 11).Value
    greatestPercentageTickerName = ws.Cells(uniqueTickertRow2, 9).Value
    greatestPercentageDecrease = ws.Cells(uniqueTickertRow2, 11).Value
    greatestDecreasePercentageTickerName = ws.Cells(uniqueTickertRow2, 9).Value
    greatestTotalTickerVolumeName = ws.Cells(uniqueTickertRow2, 9).Value
    greatestTotalTickerVolume = ws.Cells(uniqueTickertRow2, 12).Value
    uniqueTickertRow2 = uniqueTickertRow2 + 1
    'the give logic compairs the percent change and returns the greatest percent increase, greatest percent decrease and greatest total volume.
    Do Until IsEmpty(ws.Cells(uniqueTickertRow2, 11))
        If greatestPercentageIncrease < ws.Cells(uniqueTickertRow2, 11).Value Then
            greatestPercentageIncrease = ws.Cells(uniqueTickertRow2, 11).Value
            greatestPercentageTickerName = ws.Cells(uniqueTickertRow2, 9).Value
        End If
        If greatestPercentageDecrease > ws.Cells(uniqueTickertRow2, 11).Value Then
            greatestPercentageDecrease = ws.Cells(uniqueTickertRow2, 11).Value
            greatestDecreasePercentageTickerName = ws.Cells(uniqueTickertRow2, 9).Value
        End If
        If greatestTotalTickerVolume < ws.Cells(uniqueTickertRow2, 12).Value Then
            greatestTotalTickerVolume = ws.Cells(uniqueTickertRow2, 12).Value
            greatestTotalTickerVolumeName = ws.Cells(uniqueTickertRow2, 9).Value
        End If
       uniqueTickertRow2 = uniqueTickertRow2 + 1
    Loop
    'the given code tells VBA where the values are to be stored and also uses auto fit to give enough space for the data in columns
    ws.Cells(3, 15).Value = "Greatest % increase"
    ws.Columns(15).AutoFit
    ws.Cells(3, 17).Value = greatestPercentageIncrease
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(3, 16).Value = greatestPercentageTickerName
    ws.Cells(4, 15).Value = "Greatest % decrease"
    ws.Cells(4, 17).Value = greatestPercentageDecrease
    ws.Cells(4, 17).NumberFormat = "0.00%"
    ws.Cells(4, 16).Value = greatestDecreasePercentageTickerName
    ws.Cells(5, 17).Value = greatestTotalTickerVolume
    ws.Cells(5, 16).Value = greatestTotalTickerVolumeName
    ws.Cells(5, 15).Value = "Greatest Total Volume"
    ws.Columns(17).AutoFit
    ws.Columns(10).AutoFit
    ws.Columns(11).AutoFit
    ws.Columns(12).AutoFit
Next ws
End Sub
