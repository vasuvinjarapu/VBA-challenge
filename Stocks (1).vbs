Sub Stocks()

'  Loop over every Worksheet
For Each Worksheet In Worksheets

    '  Labels for the column headers
    Worksheet.Range("I1").Value = "Ticker"
    Worksheet.Range("J1").Value = "Yearly Change"
    Worksheet.Range("K1").Value = "Percent Change"
    Worksheet.Range("L1").Value = "Total Stock Volume"

    '  For Bonus Functionality Headers
    Worksheet.Range("N2").Value = "Greatest % Increase"
    Worksheet.Range("N3").Value = "Greatest % Decrease"
    Worksheet.Range("N4").Value = "Greatest Total Volume"
    Worksheet.Range("O1").Value = "Ticker"
    Worksheet.Range("P1").Value = "Value"

    ' Define data set indexes
        firstRowIndex = 2
        tickerColIndex = 1
        openColIndex = 3
        closeColIndex = 6
        volColIndex = 7
        lastRowIndex = Worksheet.Cells(Rows.Count, "A").End(xlUp).Row

    ' Define primary summary table indexes
        sum_firstRowIndex = 1
        sum_tickerColIndex = 9
        sum_yrlyChangeColIndex = 10
        sum_yrlyPercentChangeColIndex = 11
        sum_totalVolColIndex = 12

    ' Initial value for total volume summation
        totalVolume = 0

    ' Create primary summary table
        For i = firstRowIndex To lastRowIndex
                
                previousTicker = Worksheet.Cells(i - 1, 1).Value
                currentTicker = Worksheet.Cells(i, 1).Value
                nextTicker = Worksheet.Cells(i + 1, 1).Value

            '  Condtion for first instance of a ticker symbol
            If currentTicker <> previousTicker Then 
                        '  Assign yearly open price and begin summing daily volume
                        yrlyOpen = Worksheet.Cells(i, openColIndex)
                        totalVolume = totalVolume + Worksheet.Cells(i, volColIndex)

            '  Condtion for all instances of ticker symbol, except for the first and last instance  
            ElseIf previousTicker = nextTicker Then 
                        '  Continue summing daily volume
                        totalVolume = totalVolume + Worksheet.Cells(i, volColIndex)
            
            '  Condtion for last instance of a ticker symbol
            Else    
                        ' Continue summing daily volume, assign yearly close price and assign current ticker symbol' s row index for the primary summary table
                        totalVolume = totalVolume + Worksheet.Cells(i, volColIndex)
                        yrlyClose = Worksheet.Cells(i, closeColIndex)
                        sum_firstRowIndex = sum_firstRowIndex + 1

                        ' Calculate yearly change in price
                        yrlyChange = yrlyClose - yrlyOpen

                        ' Calculate yearly % change in price using if statement to prevent #DIV/0 error
                        If yrlyOpen = 0 Then
                            yrlyPercentChange = 0
                        Else
                            yrlyPercentChange = FormatPercent(yrlyChange / yrlyOpen)
                        End If

                        ' Print current ticker symbol' s info into primary summary table
                        Worksheet.Cells(sum_firstRowIndex, sum_tickerColIndex) = currentTicker
                        Worksheet.Cells(sum_firstRowIndex, sum_yrlyChangeColIndex) = yrlyChange
                        Worksheet.Cells(sum_firstRowIndex, sum_yrlyPercentChangeColIndex) = yrlyPercentChange
                        Worksheet.Cells(sum_firstRowIndex, sum_totalVolColIndex) = totalVolume

                        ' Reset total volume summation before moving to the next ticker symbol
                        totalVolume = 0
                

            End If
        
        Next i
    
    ' Add conditional formating to yrly change column
        sum_lastRowIndex = Worksheet.Cells(Rows.Count, "I").End(xlUp).Row
        
        For j = 2 To sum_lastRowIndex

            If Worksheet.Cells(j, sum_yrlyChangeColIndex) > 0 Then

                Worksheet.Cells(j, sum_yrlyChangeColIndex).Interior.Color = RGB(0, 255, 0)

            Else

                Worksheet.Cells(j, sum_yrlyChangeColIndex).Interior.Color = RGB(255, 0, 0)

            End If

        Next j


    ' BONUS FUNCTIONALITY
    ' Create secondary summary table

        ' Define ranges from primary summary table (to be used in secondary table)
        sum_tickerCol = Worksheet.Range("I:I")
        sum_yrlyPercentChangeCol = Worksheet.Range("K:K")
        sum_totalVolCol = Worksheet.Range("L:L")
    
        ' Find max/min values
        greatestPerInc = Application.WorksheetFunction.Max(sum_yrlyPercentChangeCol)
        greatestPerDec = Application.WorksheetFunction.Min(sum_yrlyPercentChangeCol)
        greatedTotalVol = Application.WorksheetFunction.Max(sum_totalVolCol)

        ' Find corresponding ticker symbol using index/match function
        greatestPerIncTicker = WorksheetFunction.Index(sum_tickerCol, WorksheetFunction.Match(greatestPerInc, sum_yrlyPercentChangeCol, 0))
        greatestPerDecTicker = WorksheetFunction.Index(sum_tickerCol, WorksheetFunction.Match(greatestPerDec, sum_yrlyPercentChangeCol, 0))
        greatedTotalVolTicker = WorksheetFunction.Index(sum_tickerCol, WorksheetFunction.Match(greatedTotalVol, sum_totalVolCol, 0))

        ' Print outputs into secondary summary table
        Worksheet.Range("O2").Value = greatestPerIncTicker
        Worksheet.Range("P2").Value = FormatPercent(greatestPerInc)

        Worksheet.Range("O3").Value = greatestPerDecTicker
        Worksheet.Range("P3").Value = FormatPercent(greatestPerDec)

        Worksheet.Range("O4").Value = greatedTotalVolTicker
        Worksheet.Range("P4").Value = greatedTotalVol

        ' Adding Border into secondary summary table
        With Range("N1:P4").Borders
            .LineStyle = xlContinuous
            .Color = vbBlack
            .Weight = xlThin
        End With

Next Worksheet

End Sub
