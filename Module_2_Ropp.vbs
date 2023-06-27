Attribute VB_Name = "Module1"
Sub StockAnalysis():

    Dim ws As Worksheet ' https://www.mrexcel.com/board/threads/vba-loop-worksheets-not-working.1147629/, referenced for how to loop through all worksheets
    
    For Each ws In Worksheets
    
        Dim i As Double
        Dim rowCounter As Double ' row counter variable for tracking row to print stock information
        Dim openPrice As Double 'variable to store open price
        Dim closePrice As Double 'variable to store close price
        Dim stockVolume As Double ' variable to store stock volume, stored as double to prevent overflow error
        Dim maxPercentIncrease As Double ' highest percent increase variable
        Dim maxIncreaseTicker As String ' ticker for stock with highest percent increase
        Dim maxPercentDecrease As Double ' highest percent decrease variable
        Dim maxDecreaseTicker As String ' ticker for stock with highest percent decrease
        Dim maxTotalVolume As Double ' highest total volume
        Dim maxVolumeTicker As String ' ticker for stock with highest total volume
        Dim lastRow As Double
        
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        rowCounter = 2
        stockVolume = 0
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
    
        For i = 2 To lastRow
        
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then ' if the ticker symbol in current cell is different than next cell then
            
                ws.Cells(rowCounter, 9).Value = ws.Cells(i, 1).Value ' print ticker symbol
                closePrice = ws.Cells(i, 6).Value ' set closing price
                ws.Cells(rowCounter, 10).Value = closePrice - openPrice ' print yearly change
                ws.Cells(rowCounter, 11).Value = FormatPercent((closePrice - openPrice) / openPrice, 2) ' calculate and print % change
                
                If (ws.Cells(rowCounter, 11).Value > maxPercentIncrease) Then ' update maxPercentIncrease if the current stock % increase is greater than the currently stored value
                
                    maxPercentIncrease = ws.Cells(rowCounter, 11).Value
                    maxIncreaseTicker = ws.Cells(rowCounter, 9).Value
                    
                ElseIf (ws.Cells(rowCounter, 11).Value < maxPercentDecrease) Then ' update maxPercentDecrease if the current stock % decrease is greater than the currently stored value
                
                    maxPercentDecrease = ws.Cells(rowCounter, 11).Value
                    maxDecreaseTicker = ws.Cells(rowCounter, 9).Value
                    
                End If
                
                If (ws.Cells(rowCounter, 10).Value < 0) Then ' format negative change cells
                
                    ws.Cells(rowCounter, 10).Interior.ColorIndex = 3
                    ws.Cells(rowCounter, 11).Interior.ColorIndex = 3
                    
                ElseIf (ws.Cells(rowCounter, 10).Value > 0) Then ' format positive change cells
                    
                    ws.Cells(rowCounter, 10).Interior.ColorIndex = 4
                    ws.Cells(rowCounter, 11).Interior.ColorIndex = 4
                    
                End If
                
                ws.Cells(rowCounter, 12).Value = stockVolume + ws.Cells(i, 7).Value ' print total stock volume
                
                If (ws.Cells(rowCounter, 12).Value > maxTotalVolume) Then ' update maxTotalVolume if total volume of current stock is higher than stored value
                
                    maxTotalVolume = ws.Cells(rowCounter, 12).Value
                    maxVolumeTicker = ws.Cells(rowCounter, 9).Value
                    
                End If
                
                stockVolume = 0 ' set stock volume to zero
                rowCounter = rowCounter + 1
            
            Else
            
                If stockVolume = 0 Then
                
                    openPrice = ws.Cells(i, 3).Value ' sets open price to opening price on first day of trading since stock volume will be zero
                    
                End If
                
                stockVolume = stockVolume + ws.Cells(i, 7).Value ' add to volume total
                
            End If
            
            
        Next i
        
        ws.Range("P2").Value = maxIncreaseTicker
        ws.Range("Q2").Value = FormatPercent(maxPercentIncrease, 2)
        ws.Range("P3").Value = maxDecreaseTicker
        ws.Range("Q3").Value = FormatPercent(maxPercentDecrease, 2)
        ws.Range("P4").Value = maxVolumeTicker
        ws.Range("Q4").Value = Format(maxTotalVolume, "Scientific") ' https://www.automateexcel.com/vba/format-numbers/#scientific, referenced on how to format cell to display as an exponent
        
        
        ws.Range("I1:Q" & lastRow).Columns.AutoFit ' https://stackoverflow.com/questions/17327037/how-to-select-a-range-of-the-second-row-to-the-last-row, referenced on how to select a range from row 1 to the last row
        
        
        ' reset all variables to evaluate next worksheet
        i = 2
        openPrice = 0
        maxPercentIncrease = 0
        maxPercentDecrease = 0
        maxTotalVolume = 0
        
        rowCounter = 2
        stockVolume = 0
        
    Next ws

End Sub
