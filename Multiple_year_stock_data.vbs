Sub Multiple_year_stock_data()
    Dim ws As Worksheet
    Dim Summary_Table_Row As Long
    Dim TotalVolume As Double
    Dim Ticker As String
    Dim openPrice As Double, closePrice As Double, YearlyChange As Double
    Dim PercentChange As Double
    Dim Lastrow As Long
    Dim RowCount As Long
    Dim MaxPercentChange As Double ' Variable to store the maximum percentage
    Dim MinPercentChange As Double ' Variable to store the minimum percentage
    Dim MaxTotalVolume As Double ' Variable to store the maximum total volume
    Dim MaxVolumeTicker As String ' Variable to store the ticker associated with the maximum total volume
    Dim MaxPercentTicker As String ' Variable to store the ticker associated with the maximum percentage
    Dim MinPercentTicker As String ' Variable to store the ticker associated with the minimum percentage
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Initialize variables
        TotalVolume = 0
        Summary_Table_Row = 2
        MaxPercentChange = 0
        MinPercentChange = 0
        MaxTotalVolume = 0
        
        ' Determine the last row in column A
        Lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
 
        
        ' Set column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("K:K").NumberFormat = "0.00%"
        ws.Range("L1").Value = "TotalStockVolume"
        
        ' Get the openPrice for the first Ticker
        openPrice = ws.Cells(2, 3).Value
        
        ' Loop through each row
        For i = 2 To Lastrow
       
                If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                    ' Calculate total volume for the Ticker
                    Ticker = ws.Cells(i, 1).Value
                    TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                    
                    ' Write Ticker to summary table
                    ws.Range("I" & Summary_Table_Row).Value = Ticker
                
                    
                    ' Calculate YearlyChange and PercentChange
                    closePrice = ws.Cells(i, 6).Value
                    YearlyChange = closePrice - openPrice
                    If openPrice <> 0 Then
                        PercentChange = (YearlyChange / openPrice)
                    Else
                        PercentChange = 0
                    End If
                    
                    ' Write YearlyChange,PercentChange and TotalStockVolume to summary table
                    ws.Range("J" & Summary_Table_Row).Value = YearlyChange
                    ' Apply conditional formatting on YearlyChange
                    If YearlyChange > 0 Then
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4 ' Green
                    Else
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3 ' Red
                    End If
                    ws.Range("K" & Summary_Table_Row).Value = PercentChange
                     ' Apply conditional formatting on PercentChange
                    If PercentChange > 0 Then
                    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                    Else
                      ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                    End If
                        ws.Range("L" & Summary_Table_Row).Value = TotalVolume
                    
                    ' Track the maximum percentage, minimum percentage, and maximum total volume
                    If PercentChange > MaxPercentChange Then
                        MaxPercentChange = PercentChange
                        MaxPercentTicker = Ticker
                    End If
                    
                    If PercentChange < MinPercentChange Then
                        MinPercentChange = PercentChange
                        MinPercentTicker = Ticker
                    End If
                    
                    If TotalVolume > MaxTotalVolume Then
                        MaxTotalVolume = TotalVolume
                        MaxVolumeTicker = Ticker
                    End If
                    
                    ' Reset variables for the next Ticker
                    TotalVolume = 0
                    openPrice = ws.Cells(i + 1, 3).Value
                    Summary_Table_Row = Summary_Table_Row + 1
                Else
                    ' Accumulate TotalVolume for the Ticker
                    TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                End If
      
        Next i
        
        ' Write summary information for each sheet
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P2").Value = MaxPercentTicker
        ws.Range("Q2").Value = Format(MaxPercentChange, "0.00%")
        ws.Range("P3").Value = MinPercentTicker
        ws.Range("Q3").Value = Format(MinPercentChange, "0.00%")
        ws.Range("P4").Value = MaxVolumeTicker
        ws.Range("Q4").Value = MaxTotalVolume
        
   
    Next ws
End Sub

