Attribute VB_Name = "Module1"
Sub StockDisplay():

    Dim WorkSheetCount As Integer
    Dim i As Integer
    
        
    
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim NumPercent As Double
    Dim Change As Double
    Dim StockVolume As Double
    Dim TickerSymbol As String 'string for ticker symbol letters
    Dim Row As Long 'Long since there are many many rows
    Dim PrintRow As Long 'Long since there are many many rows
    'Dim Clmn As Range
    'Dim posCondition As FormatCondition, negCondition As FormatCondition

    Dim GreatestIncrease As Double
    Dim GreatestIncreaseTS As String
    Dim GreatestDecrease As Double
    Dim GreatestDecreaseTS As String
    Dim GreatestTV As Double
    Dim GreatestTVTS As String
    Dim TickerSearch As String
    Dim TSRow As Long
    
    

    'Initialize variable and loop through all sheets in workbook
    WorkSheetCount = ActiveWorkbook.Worksheets.Count

    For i = 1 To WorkSheetCount

        Worksheets(i).Activate
        
        'Initialize Variables to Test Against
        
        TickerSymbol = Cells(2, 1).Value
        Row = 2
        OpeningPrice = Cells(Row, 3).Value
        ClosingPrice = 0
        NumPercent = 0
        StockVolume = 0
        PrintRow = 2
        
        'Print Row Headers
        Cells(1, 10) = "Ticket Symbol"
        Cells(1, 11) = "Change in Price"
        Cells(1, 12) = "Percent Change"
        Cells(1, 13) = "Total Stock Volume"
        
        
        'Run through lines of data until an empty row is found, then stop
        
        While TickerSymbol <> Empty
            '***THIS ANALYSIS ASSUMES INFORMATION PER
            '***TICKER SYMBOL IS IN CHRONOLOGICAL ORDER
            
            'if ticker symbol is the same as the last ticker symbol
            'increase stock volume and update closing price to new latest closing price
            If TickerSymbol = Cells(Row, 1).Value Then
                StockVolume = StockVolume + Cells(Row, 7).Value
                ClosingPrice = Cells(Row, 6)
                
            'if new ticker symbol if found, print data for previous symbol
            Else
                'print ticker symbol of current row
                Cells(PrintRow, 10) = TickerSymbol
                
                'print closing price of current row
                Change = (ClosingPrice - OpeningPrice)
                Cells(PrintRow, 11) = Change
                
                'Calculate and print % change from initialized opening
                If OpeningPrice <> 0 Then
                    NumPercent = ((ClosingPrice - OpeningPrice) / OpeningPrice)
                Else
                    NumPercent = 0
                End If
                    
                Cells(PrintRow, 12).Value = NumPercent
                
                'print totaled stock volume
                Cells(PrintRow, 13).Value = StockVolume
                
                'reset values for new ticker symbol
                StockVolume = 0
                NumPercent = 0
                OpeningPrice = Cells(Row, 3)
                TickerSymbol = Cells(Row, 1).Value
                
                'increase print row for next print
                PrintRow = PrintRow + 1
                
                            
            End If
            
            'increase row for data collection
            Row = Row + 1
            
        Wend
        
        Columns("L").NumberFormat = "0.00%"
        
        'Set Clmn = Range("K2:K2000")
        
        'Clmn.FormatConditions.Delete
        
        'Set posCondition = Clmn.FormatConditions.Add(xlColorScale, xlGreater, "0")
    
        'Set negCondition = Clmn.FormatConditions.Add(xlColorScale, xlLess, "0")
        
        'With posCondition
        '    .Interior.Color = vbGreen
         '   End With
            
       ' With negCondition
        '    .Interior.Color = vbRed
        '    End With
        
        
        'Use if statement to format colors
        Row = 2
        Change = Cells(Row, 11).Value
        TickerSymbol = Cells(Row, 10).Value
        
        While TickerSymbol <> Empty
    
            If Change < 0 Then
                Cells(Row, 11).Interior.ColorIndex = 3 'Red
                
            Else 'zero taken as positive
                Cells(Row, 11).Interior.ColorIndex = 4 'Green
                
            End If
            
            Row = Row + 1
            Change = Cells(Row, 11).Value
            TickerSymbol = Cells(Row, 10).Value
            
            Wend


        'Challenges
        
        TSRow = 2
        TickerSearch = Cells(TSRow, 10)
        GreatestIncrease = Cells(TSRow, 12).Value
        GreatestIncreaseTS = Cells(TSRow, 10)
        GreatestDecrease = Cells(TSRow, 12).Value
        GreatestDecreaseTS = Cells(TSRow, 10)
        GreatestTV = Cells(TSRow, 13).Value
        GreatestTVTS = Cells(TSRow, 10)
        
        While TickerSearch <> Empty
            TSRow = TSRow + 1
            
            If GreatestIncrease < Cells(TSRow, 12).Value Then
                GreatestIncrease = Cells(TSRow, 12).Value
                GreatestIncreaseTS = Cells(TSRow, 10)
            End If
            
            If GreatestDecrease > Cells(TSRow, 12).Value Then
                GreatestDecrease = Cells(TSRow, 12).Value
                GreatestDecreaseTS = Cells(TSRow, 10)
            End If
            
            If GreatestTV < Cells(TSRow, 13).Value Then
                GreatestTV = Cells(TSRow, 13).Value
                GreatestTVTS = Cells(TSRow, 10)
            End If
            
            TickerSearch = Cells(TSRow, 10)
            
            
        Wend
        
        Range("O2") = "Greatest % Increase"
        Range("O3") = "Greatest % Decrease"
        Range("O4") = "Greatest % Total Volume"
        Range("P1") = "Ticker"
        Range("Q1") = "Value"
        Range("P2") = GreatestIncreaseTS
        Range("P3") = GreatestDecreaseTS
        Range("P4") = GreatestTVTS
        Range("Q2") = GreatestIncrease
        Range("Q3") = GreatestDecrease
        Range("Q2").NumberFormat = "0.00%"
        Range("Q3").NumberFormat = "0.00%"
        Range("Q4") = GreatestTV

    Next i





End Sub



