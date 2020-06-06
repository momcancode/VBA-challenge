Attribute VB_Name = "Module1"
Sub WallStreet()

' Loop through each sheet in the workbook
For Each ws In Worksheets

    ' Set variables holding ticker symbol, closing price, opening price,
    ' yearly change, and percent change
    ' Declare variable types
    Dim Ticker As String
    Dim ClosingPrice As Double
    Dim OpeningPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    
    ' Keep track of locations of each ticker symbol in the First Summary Table
    Dim SummaryRow As Long
    ' Start the first location at the second row
    SummaryRow = 2
    
    ' Set a variable holding the number of rows per ticker block in the Original Stock Data table
    ' Except the row holding the last trading day of each ticker in a given year
    Dim TickerCount As Long
    TickerCount = 0
      
    ' Set a variable holding the cumulative stock volume
    Dim cumulativeVolume As Double
    cumulativeVolume = 0
   
    ' Set a variable for the row number holding opening price of each ticker at a given year
    Dim OpeningPriceRow As Long
            
    ' Determine the last row of the Original Stock Data Table
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Delete rows with no meaningful data in Original Stock Data table
    ' (i.e. rows with no opening prices, closing prices, and stock volumes)
    ' Loop through all data rows in Original Stock Data Table
    For j = lastRow To 2 Step -1
        If ws.Cells(j, 3).Value = 0 And ws.Cells(j, 6).Value = 0 And ws.Cells(j, 7).Value = 0 Then
            ws.Rows(j).EntireRow.Delete
        End If
    Next j
    
    ' Sort <ticker> and <date> columns to make sure they are in ascending order
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("A1"), Order:=xlAscending
        .SortFields.Add Key:=Range("B1"), Order:=xlAscending
        .SetRange Range("A1:G" & lastRow)
        .Header = xlYes
        .Apply
    End With
       
    ' Loop through all data rows in Original Stock Data Table
    For r = 2 To lastRow
        
        ' Check if we are still within the same ticker symbol. If not then ...
        If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
            
            ' Get the ticker symbol and output to the First Summary Table
            Ticker = ws.Cells(r, 1).Value
            ws.Cells(SummaryRow, 9).Value = Ticker
            
            ' Get the closing price at the end of the year
            ClosingPrice = ws.Cells(r, 6).Value

            ' Find the row number holding each ticker at the beginning of a year by
            ' subtracting the TickerCount from the row number holding the ticker
            ' at the end of the year
            OpeningPriceRow = r - TickerCount
            
            ' Find opening price at the beginning of a year
            OpeningPrice = ws.Cells(OpeningPriceRow, 3).Value
            
            ' Calculate the yearly change from opening price at the beginning of a given year
            ' to the closing price at the end of that year
            YearlyChange = ClosingPrice - OpeningPrice
            
            ' Print the yearly change to the First Summary Table
            ws.Cells(SummaryRow, 10).Value = YearlyChange
            
            ' Format the Yearly Change column to Number
            ws.Cells(SummaryRow, 10).NumberFormat = "0.00"
            
            ' Conditional formatting that will highlight positive change in green
            ' and negative change in red
            If YearlyChange > 0 Then
                ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
            End If
            
            ' Calculate Percent Change
            PercentChange = YearlyChange / OpeningPrice
            
            ' Print the Percent Change to the First Summary Table
            ws.Cells(SummaryRow, 11).Value = PercentChange
            
            ' Format the Percent Change column to Percentage
            ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
            
            ' Tally up the stock volumes as we iterate
            cumulativeVolume = cumulativeVolume + ws.Cells(r, 7).Value
            
            ' Print the Total Stock Volume to the First Summary Table
            ws.Cells(SummaryRow, 12).Value = cumulativeVolume
            
            ' Reset TickerCount and cumulativeVolume
            TickerCount = 0
            cumulativeVolume = 0
            
            ' Add one to the summary table row before looping through rows under a new Ticker
            SummaryRow = SummaryRow + 1

        ' If the cell immediately following a row is the same ticker...
        Else
        
            ' Add one to the TickerCount as we iterate
            TickerCount = TickerCount + 1
            
            ' Tally up the stock volumes as we iterate
            cumulativeVolume = cumulativeVolume + ws.Cells(r, 7).Value
                        
        End If
    
    Next r
    
    ' Label headers of columes I to L of the First Summary Table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    '------------------------------------------------------------------------------------------
    ' CHALLENGES
    
    ' Determine the Last Row in the First Summary Table
    LastSummaryTableRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    ' Find the Greatest % increase, the Greatest % decrease and
    ' the Greatest total volume in the First Summary Table
    MaxPercentChange = Application.WorksheetFunction.Max(ws.Range("K2:K" & LastSummaryTableRow))
    MinPercentChange = Application.WorksheetFunction.Min(ws.Range("K2:K" & LastSummaryTableRow))
    MaxVolume = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastSummaryTableRow))
    
    ' Loop through data rows in the First Summary Table
    For i = 2 To LastSummaryTableRow
    
        ' Get ticker, percent change and total stock volume as iterating through the summary table
        SummaryTicker = ws.Cells(i, 9).Value
        SummaryPercentChange = ws.Cells(i, 11).Value
        SummaryTotalVolume = ws.Cells(i, 12).Value
        
        ' Check if the percent change is equal to the Greatest % increase. If yes then ...
        If SummaryPercentChange = MaxPercentChange Then
            ' Print the ticker and the percent change to the Second Summary Table
            ws.Range("P2").Value = SummaryTicker
            ws.Range("Q2").Value = SummaryPercentChange
        End If
        
        ' Check if the percent change is equal to the Greatest % decrease. If yes then ...
        If SummaryPercentChange = MinPercentChange Then
            ' Print the ticker and the percent change to the Second Summary Table
            ws.Range("P3").Value = SummaryTicker
            ws.Range("Q3").Value = SummaryPercentChange
        End If
        
        ' Check if the total volume is equal to the greatest total volume. If yes then ...
        If SummaryTotalVolume = MaxVolume Then
            ' Print the ticker and the total volume to the Second Summary Table
            ws.Range("P4").Value = SummaryTicker
            ws.Range("Q4").Value = SummaryTotalVolume
        End If
    Next i
    
    ' Format the Greastest % increase and decrease to Percentage
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    ' Label the headers of rows and columns in the Second Summary Table
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    ' Autofit to display data
    ws.Columns("A:Q").AutoFit
    
    
Next ws

End Sub
