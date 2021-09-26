Attribute VB_Name = "Module1"
Sub WallStreet():
    
    ' define variables used for analysis
    Dim tickerSymbol As String
    
    Dim volumeTotal As Long
    volumeTotal = 0

    Dim lastRow As Long
    
    Dim openPrice As Double
    openPrice = 0
    
    Dim closePrice As Double
    closePrice = 0
    
    Dim yearlyPriceChange As Double
    yearlyPriceChange = 0
    
    Dim percentPriceChange As Double
    percentPriceChange = 0
    
    Dim maxPrecent As Double
    maxPercent = 0
    
    Dim minPercent As Double
    minPercent = 0
    
    Dim maxTicker As String
    
    Dim minTicker As String
    
    Dim summaryTableRow As Long
    summaryTableRow = 2 ' starts at row 2 in the summary table
    
    ' count the number of rows
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' loop through all the ticker rows
    For Row = 2 To lastRow
    
        ' check to see if we are still in the same ticker symbol
        If (Cells(Row, 1).Value <> Cells(Row + 1, 1).Value) Then
        
            ' set the ticker symbol
            tickerSymbol = Range("A" & Row).Value
            
            ' calculate
            closePrice = Range("F" & Row).Value
            yearlyPriceChange = closePrice - openPrice
            
            ' set the condition for zero
            If openPrice <> 0 Then
                percentPriceChange = (yearlyPriceChange / openPrice) * 100
                
            End If
            
            ' next opening price
            openPrice = Cells(Row + 1, 3).Value
            
            ' calulations for percent change
            If (percentPriceChange > maxPercent) Then
                maxPercent = percentPriceChange
                maxTicker = tickerSymbol
                
            ElseIf percentPriceChange < minPercent Then
                minPercent = percentPriceChange
                minTicker = tickerSymbol
                
            End If
            
            ' add to the volume total one last time before the change in ticker
            volumeTotal = volumeTotal + Range("G" & Row).Value
            
            ' add the values to the summary table
            
            ' add the ticker symbol to Column I on the current summary table row
            Range("I" & summaryTableRow).Value = tickerSymbol
            
            ' add final volume total to Column L on the current summary table row
            Range("L" & summaryTableRow).Value = volumeTotal
            
            ' add yearly change to Column J on the current summay table row
            Range("J" & summaryTableRow).Value = yearlyPriceChange
            
            If yearlyPriceChange > 0 Then
                Range("J" & summaryTableRow).Interior.ColorIndex = 4
                
            ElseIf yearlyPriceChange <= 0 Then
                Range("J" & summaryTableRow).Interior.ColorIndex = 3
            
            End If
            
            ' add percent change to Column K on the current summary table row
            Range("K" & summaryTableRow).Value = (CStr(percentPriceChange) & "%")
            
            ' once the summary table is populated, then add one to the summary row count
            summaryTableRow = summaryTableRow + 1
            
            'then reset the volume total to 0
            volumeTotal = 0
            
            ' if we are in the same ticker symbol, add on to the running total
            volumeTotal = volumeTotal + Range("G" & Row).Value
        
            ' color fill price change for red is negative and green is positiver
            
        End If
    
    Next Row
    
End Sub

