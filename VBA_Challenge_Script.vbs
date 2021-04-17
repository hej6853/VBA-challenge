Option Explicit

Sub TickerCalculation():

Dim dataRow As Long
Dim outputRow  As Long
Dim PercentChange As Long
Dim SheetNum As Long
Dim RowCount As Long
Dim row_number_greatestincrease As Long
Dim row_number_greatestdecrease As Long
Dim row_number_greatestttlvolume As Long


' I know this is first ticker, first row
' therefore, save the open price
' Create a counter for total stock volume

Dim openPrice As Double
Dim totalStockVolume As Double
Dim closePrice As Double

For SheetNum = 1 To Worksheets.Count
    Dim ws As Worksheet
    Set ws = Worksheets(SheetNum)
    outputRow = 2
  
  'Output
    'Col 10 - Ticker
    ws.Range("I1").Value = "Ticker"
    'Col 11 - Yearly Change
    ws.Range("J1").Value = "Yearly Change"
    'Col 12 - Percent Change
    ws.Range("K1").Value = "Percent Change"
    'Col 13 - Total Stock Volume
    ws.Range("L1").Value = "Total Stock Volume"
    'Col 15 - Greatest % Increase
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    'Col 16 - Ticker
    ws.Range("P1").Value = "Ticker"
    'Col 17 - Value
    ws.Range("Q1").Value = "Value"



    openPrice = ws.Range("C2").Value
    ' Start loop at A2
    For dataRow = 2 To ws.Range("A2").End(xlDown).Row
            If ws.Cells(dataRow, 1).Value <> ws.Cells(dataRow + 1, 1).Value Then
            ' Now at the edge
            ' add whatever is in Col to G to the total stock volume counter
            totalStockVolume = (totalStockVolume + ws.Cells(dataRow, 7).Value)
            ' grab the closing price from Col F
            closePrice = ws.Cells(dataRow, 6).Value
            ' Now calculate the yearly change as closeprice - openprice
            ' Calculate yearly percent change as closeprice - openprice/ openprice
            ' Since there might be a division by 0, put in check to make sure that the denominator is not 0
            ' Copy over the value in Col A to Col I
            ' Then dump the yearly change, percen change and total stock volume into J,K,L

            'Percent Change
            If openPrice = 0 Then
                ws.Cells(outputRow, 11).Value = "NaN"
            Else
                ws.Cells(outputRow, 11).Value = (closePrice - openPrice) / openPrice
                ws.Cells(outputRow, 11).Value = Format(ws.Cells(outputRow, 11).Value, "#.##%")
    
            End If

            'Yearly Change
            ws.Cells(outputRow, 10).Value = closePrice - openPrice
            If openPrice > closePrice Then
                ws.Cells(outputRow, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(outputRow, 10).Interior.ColorIndex = 4
            End If

            'Total Stock Volume
            ws.Cells(outputRow, 12).Value = totalStockVolume

            'Ticker
            ws.Cells(outputRow, 9).Value = ws.Cells(dataRow, 1).Value

            ' Add 1 to the row counter for the output table
            outputRow = outputRow + 1
            
            ' Then update the new open price to be the open price of the next row
            totalStockVolume = 0
            openPrice = ws.Cells(dataRow + 1, 3).Value
        Else
            ' If it's not the edge, then
            ' don't change the open value
            ' add whatever is in Col G to the total stock volume' counter
            totalStockVolume = (totalStockVolume + ws.Cells(dataRow, 8).Value)

        End If
    Next dataRow

    ' Greatest % Increase
    RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & RowCount)) * 100

    row_number_greatestincrease = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)

     ws.Range("P2") = ws.Cells(row_number_greatestincrease + 1, 9).Value
   
      

     'Greatest % Decrease
     RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
     ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & RowCount)) * 100
     row_number_greatestdecrease = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
     ws.Range("P3") = ws.Cells(row_number_greatestdecrease + 1, 9).Value

     'Greatest Total Volume
     RowCount = Cells(Rows.Count, "A").End(xlUp).Row
     ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & RowCount))
     row_number_greatestttlvolume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & RowCount)), ws.Range("L2:L" & RowCount), 0)
     ws.Range("P4") = ws.Cells(row_number_greatestttlvolume + 1, 9).Value
    
Next SheetNum
End Sub










