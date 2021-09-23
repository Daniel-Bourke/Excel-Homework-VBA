Attribute VB_Name = "Module1"
Sub StockAnalyis()
    'Setting Names for Headers and Metrics
    NewHeaders = Array("Ticker", "Opening", "Closing", "Yearly Change", "% Change", "Total Stock Volume", "Metric", "StockID", "Value")
    Metrics = Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume")
    
    For Each ws In Worksheets
        
        Dim TickerRow As Long
        Dim LastTickerRow As Long
        
        Dim SummaryTableRow As Integer
        Dim LastSummaryTableRow As Integer
        
        Dim StockID As String
        Dim CumulativeStockVolume As Double
        Dim OpeningStockPrice As Double
        Dim ClosingStockPrice As Double
        Dim YearlyChange As Double
        Dim PercentYearlyChange As Double
        
        'Set Out the Layout
        'Label the Summary Table headers
        ws.Range("I1:Q1").Value = NewHeaders
        
        'Label the metric names
        'transpose the metrics and write them into the metric table
        ws.Range("O2:O4").Value = Application.WorksheetFunction.Transpose(Metrics)
        
        'Initial Values
        CumulativeStockVolume = 0
        'The summary table data begins on row 2
        SummaryTableRow = 2
        
        'Setting initial OpeningStockPrice for the first Stock
        OpeningStockPrice = ws.Cells(2, 3).Value
        'Write the OpeningStockPrice for the first Stock in the summary table
        ws.Range("j" & SummaryTableRow).Value = OpeningStockPrice
        
        'Count the number of rows in data
        LastTickerRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'go through each row in the data
        For TickerRow = 2 To LastTickerRow
            
            'If the StockID of the current row is different to the StockID of the row below the current row
            If ws.Cells(TickerRow + 1, 1).Value <> ws.Cells(TickerRow, 1).Value Then
                
                'Get the StockID
                StockID = ws.Cells(TickerRow, 1).Value
                'Write the StockID in the summary table
                ws.Range("i" & SummaryTableRow).Value = StockID
                
                'Get the CumulativeStockVolume so far by adding this rows stock volume to the previous cumulative
                CumulativeStockVolume = CumulativeStockVolume + ws.Cells(TickerRow, 7).Value
                'Write the CumulativeStockVolume for the current StockID in the summary table
                ws.Range("n" & SummaryTableRow).Value = CumulativeStockVolume
                
                'Get the ClosingStockPrice for this Stock
                ClosingStockPrice = ws.Cells(TickerRow, 6).Value
                'Write the ClosingStockPrice in the summary table
                ws.Range("k" & SummaryTableRow).Value = ClosingStockPrice
                'Calculate YearlyChange
                YearlyChange = (ClosingStockPrice - OpeningStockPrice)
                'Write YearlyChange for each ticker in the summary table
                ws.Range("l" & SummaryTableRow).Value = YearlyChange
                
                'Get the PercentYearlyChange for this stock (account for error when denominator is 0)
                If OpeningStockPrice = 0 Then
                    PercentYearlyChange = 0
                Else
                    PercentYearlyChange = (ClosingStockPrice - OpeningStockPrice) / OpeningStockPrice
                End If
                
                'Write the PercentYearlyChange for each ticker in the summary table
                ws.Range("m" & SummaryTableRow).Value = PercentYearlyChange
                ws.Range("m" & SummaryTableRow).NumberFormat = "0.00%"
                
                'Move on to the next row in the summary table
                SummaryTableRow = SummaryTableRow + 1
                
                'Reset CumulativeStockVolume
                CumulativeStockVolume = 0
                
                'get the OpeningStockPrice of the next Stock
                OpeningStockPrice = ws.Cells(TickerRow + 1, 3)
                'Write the OpeningOpeningStockPrice of the Next stock in the Summary table
                ws.Range("j" & SummaryTableRow).Value = OpeningStockPrice
                
            Else
                'Because it's the same stock we want to add this rows volume to the CumulativeStockVolume
                CumulativeStockVolume = CumulativeStockVolume + ws.Cells(TickerRow, 7).Value
                
            End If
            
        Next TickerRow
        
        'For each SummaryTableRow
        LastSummaryTableRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        For SummaryTableRow = 2 To LastSummaryTableRow
            'if the YearlyChange was Positive then green fill, otherwise red fill in summary table
            If ws.Cells(SummaryTableRow, 12).Value > 0 Then
                ws.Cells(SummaryTableRow, 12).Interior.ColorIndex = 51
                
            Else
                ws.Cells(SummaryTableRow, 12).Interior.ColorIndex = 30
                
            End If
            
        Next SummaryTableRow
        'Change Font of Yearly Change to White and Bold
        ws.Range("l2:l" & LastSummaryTableRow).Font.Bold = True
        ws.Range("l2:l" & LastSummaryTableRow).Font.Color = vbWhite
        
        'Determine the "Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume"
        
        '
        For SummaryTableRow = 2 To LastSummaryTableRow
            
            'Find the Greatest % Increase (Max of the column m)
            If ws.Cells(SummaryTableRow, 13).Value = Application.WorksheetFunction.Max(ws.Range("m2:m" & LastSummaryTableRow)) Then
                'Write the ticker name
                ws.Cells(2, 16).Value = ws.Cells(SummaryTableRow, 9).Value
                'Write the Greatest % Increase
                ws.Cells(2, 17).Value = ws.Cells(SummaryTableRow, 13).Value
                'Format as %
                ws.Cells(2, 17).NumberFormat = "0.00%"
                
                'Find the Greatest % Decrease (Min of the column m)
            ElseIf ws.Cells(SummaryTableRow, 13).Value = Application.WorksheetFunction.Min(ws.Range("m2:m" & LastSummaryTableRow)) Then
                'Write the ticker name
                ws.Cells(3, 16).Value = ws.Cells(SummaryTableRow, 9).Value
                'Write the Greatest % Decrease
                ws.Cells(3, 17).Value = ws.Cells(SummaryTableRow, 13).Value
                'Format as %
                ws.Cells(3, 17).NumberFormat = "0.00%"
                
                'Find the Greatest Total Volume (Max of the column n)
            ElseIf ws.Cells(SummaryTableRow, 14).Value = Application.WorksheetFunction.Max(ws.Range("n2:n" & LastSummaryTableRow)) Then
                'Write the ticker name
                ws.Cells(4, 16).Value = ws.Cells(SummaryTableRow, 9).Value
                'Write the Greatest Total Volume
                ws.Cells(4, 17).Value = ws.Cells(SummaryTableRow, 14).Value
                
            End If
            
        Next SummaryTableRow
        
    Next ws
    
End Sub
