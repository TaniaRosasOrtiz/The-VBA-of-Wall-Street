Attribute VB_Name = "Module1"
Sub Stock_Market_Analysis()
        
'  DECLARE VARIABLES
    Dim lastrow, tickerIndex, rowTickers, ws_count, I As Integer
    Dim ticker, addressRange As String
    Dim openPrice, closePrice, yearChangePrice, yearChangePricePerc, totalStock As Double
    'BONUS VAR:
        Dim maxIncreaseCell, maxDecreaseCell, maxVolumeCell, summaryRange As Range

' IDENTIFY ALL SHEETS
    ws_count = ActiveWorkbook.Worksheets.Count
        
    For I = 1 To ws_count
    
        ActiveWorkbook.Worksheets(I).Activate
    
        ' INITIALIZE VARIABLES
            lastrow = 2
            yearChangePrice = 0
            yearChangePricePerc = 0
            tickerIndex = 0
            rowTickers = 2
            
            'GET FIRST TICKER NAME
                ticker = Cells(lastrow, 1).Value
        
        ' LABEL HEADERS FOR RESULTS
            Cells(1, 9).Value = "Ticker"
            Cells(1, 10).Value = "Yearly Change"
            Cells(1, 11).Value = "Percent Change"
            Cells(1, 12).Value = "Total Stock Volume"
        
        'START READING VALUES FOR THE SAME TICKER
            While ticker <> ""
            
                totalStock = totalStock + Cells(lastrow, 7).Value
                tickerIndex = tickerIndex + 1
                
            'GET openPrice AT THE BEGINING OF THE YEAR
                If tickerIndex = 1 Then
                    openPrice = Cells(lastrow, 3).Value
                End If
                
            'GET closePrice AT THE END OF THE YEAR FOR THE TICKER
                If ticker <> Cells(lastrow + 1, 1).Value Then
                    
                    closePrice = Cells(lastrow, 6).Value
                    
                    'COMPUTE ALL VARIABLES FOR SAME TICKER
                        yearChangePrice = closePrice - openPrice
                        
                        If openPrice = 0 Then
                            yearChangePricePerc = 0
                            Else
                                yearChangePricePerc = (yearChangePrice / openPrice)
                        End If
                                
                    'WRITE RESULTS
                        Cells(rowTickers, 9).Value = ticker
                        Cells(rowTickers, 10).Value = yearChangePrice
                        Cells(rowTickers, 11).Value = yearChangePricePerc
                        Cells(rowTickers, 11).NumberFormat = "0.00%"
                        Cells(rowTickers, 12).Value = totalStock
                    
                    'FORMAT yearChangePrice
                        If yearChangePrice < 0 Then
                            Cells(rowTickers, 10).Interior.ColorIndex = 3 'RED FOR yearChangePrice NEGATIVE
                        Else
                            Cells(rowTickers, 10).Interior.ColorIndex = 4 'GREEN FOR yearChangePrice POSITIVE
                        End If
                    
                'RESET VALUES FOR NEW TICKER
                    totalStock = 0
                    tickerIndex = 0
                    rowTickers = rowTickers + 1
                    
                End If
                
                lastrow = lastrow + 1
                
            'GET NEW TICKER
                ticker = Cells(lastrow, 1).Value
           
            Wend
        
        ' BONUS: LABEL HEADERS FOR TOP SOLUTIONS
            Cells(1, 15).Value = "Ticker"
            Cells(1, 16).Value = "Value"
            Cells(2, 14).Value = "Greatest % Increase"
            Cells(3, 14).Value = "Greatest % Decrease"
            Cells(4, 14).Value = "Greatest Total Volume"
            
            rowTickers = rowTickers - 1
            
            
            ' CALCULATE maxVolumeCell
                addressRange = "L2:L" + CStr(rowTickers)
                Set summaryRange = Range(addressRange)
                maxVolumeCell = AddressOfMax(summaryRange).Address
                Cells(4, 16).Value = Range(maxVolumeCell).Value
                Range(maxVolumeCell).Activate
                ActiveCell.Offset(, columnOffset:=-3).Activate
                ticker = ActiveCell.Value
                Cells(4, 15).Value = ticker
            
            ' CALCULATE maxIncreaseCell
                addressRange = "K2:K" + CStr(rowTickers)
                Set summaryRange = Range(addressRange)
                maxIncreaseCell = AddressOfMax(summaryRange).Address
                Cells(2, 16).Value = Range(maxIncreaseCell).Value
                Cells(2, 16).NumberFormat = "0.00%"
                Range(maxIncreaseCell).Activate
                ActiveCell.Offset(, columnOffset:=-2).Activate
                ticker = ActiveCell.Value
                Cells(2, 15).Value = ticker
                
            ' CALCULATE maxDecreaseCell (USING SAME COLUMN, THEREFORE summaryRange DO NOT CHANGE)
                maxDecreaseCell = AddressOfMin(summaryRange).Address
                Cells(3, 16).Value = Range(maxDecreaseCell).Value
                Cells(3, 16).NumberFormat = "0.00%"
                Range(maxDecreaseCell).Activate
                ActiveCell.Offset(, columnOffset:=-2).Activate
                ticker = ActiveCell.Value
                Cells(3, 15).Value = ticker
    
        Range("A1").Activate
    
    Next I   'CHANGE SHEET
    
End Sub

Function AddressOfMax(rng As Range) As Range
    Set AddressOfMax = rng.Cells(WorksheetFunction.Match(WorksheetFunction.Max(rng), rng, 0))
End Function
Function AddressOfMin(rng As Range) As Range
    Set AddressOfMin = rng.Cells(WorksheetFunction.Match(WorksheetFunction.Min(rng), rng, 0))
End Function
