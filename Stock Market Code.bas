Attribute VB_Name = "Module1"
Sub Stonks()
    'Declaring variables
    Dim LastColumn As Double
    Dim LastRow As Double
    Dim Volume As Double
    Dim YearlyChange As Double
    Dim PctChange As Double
    Dim FirstOfTick As Double
    Dim Ticker As String
    Dim StockNumber As Double
    Dim LastTick As Double
    Dim GreatestPerIn As Double
    Dim GreatestPerDec As Double
    Dim GreatestVol As Double
    Dim GreatestPerInName As String
    Dim GreatestPerDecName As String
    Dim GreatestVolName As String
    Dim ws As Worksheet
    Dim i As Double
    Dim j As Double
    For Each ws In ActiveWorkbook.Worksheets
    
        'First iteration for final numbers
        i = 2
        StockNumber = 2
        FirstOfTick = 2
        'Getting last row and column for for loop and placing final values, respectively
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        LastColumn = ws.Cells(LastRow, ws.Columns.Count).End(xlToLeft).Column
        'Naming final columns and indexing for future use
        ws.Cells(1, LastColumn + 2).Value = "Ticker"
        TickerC = LastColumn + 2
        ws.Cells(1, LastColumn + 3).Value = "Yearly Change"
        YearC = LastColumn + 3
        ws.Cells(1, LastColumn + 4).Value = "Percent Change"
        PercentC = LastColumn + 4
        ws.Cells(1, LastColumn + 5).Value = "Total Stock Volume"
        VolumeC = LastColumn + 5
        'Iterating across the rows to pull the needed data
        For i = 2 To LastRow
            'Finds the point where the ticker changes to the next stock
            If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
                'Grabs the ticker name, yearly change, percent change, and sum of volume
                Ticker = ws.Cells(i, 1).Value
                YearlyChange = ws.Cells(i, 6).Value - Cells(FirstOfTick, 3).Value
                PctChange = YearlyChange / ws.Cells(FirstOfTick, 3).Value
                Volume = WorksheetFunction.Sum(ws.Range("G" & FirstOfTick & ":G" & i))
                'Sets the cells associated with the given ticker to the required values
                ws.Cells(StockNumber, TickerC).Value = Ticker
                ws.Cells(StockNumber, YearC).Value = YearlyChange
                ws.Cells(StockNumber, PercentC).Value = PctChange
                ws.Cells(StockNumber, VolumeC).Value = Volume
                'Defines the first row of the next data set
                FirstOfTick = i + 1
                'Increases the row index by one for the final data
                StockNumber = StockNumber + 1
            End If
        Next i
    
        'Identifying the last row of the final data
        LastTick = ws.Cells(ws.Rows.Count, (LastColumn + 2)).End(xlUp).Row
        GreatestPerIn = 0
        GreatestPerDec = 0
        GreatestVol = 0
        For j = 2 To LastTick
            'Finding the largest percent increase
            If ws.Cells(j, PercentC).Value > GreatestPerIn Then
                GreatestPerIn = ws.Cells(j, PercentC).Value
                GreatestPerInName = ws.Cells(j, TickerC).Value
            End If
            'Finding the largest percent decrease
            If ws.Cells(j, PercentC).Value < GreatestPerDec Then
                GreatestPerDec = ws.Cells(j, PercentC).Value
                GreatestPerDecName = ws.Cells(j, TickerC).Value
            End If
            'Finding the largest total volume
            If ws.Cells(j, VolumeC).Value > GreatestVol Then
                GreatestVol = ws.Cells(j, VolumeC).Value
                GreatestVolName = ws.Cells(j, TickerC).Value
            End If
            'Setting the color of the cell based on stock performance
            If ws.Cells(j, YearC).Value < 0 Then
                ws.Cells(j, YearC).Interior.ColorIndex = 3
            Else
            ws.Cells(j, YearC).Interior.ColorIndex = 4
            End If
        Next j
        'Creating the labels for the greatest values
        Dim Labels As Integer
        Labels = VolumeC + 3
        With ws
            .Cells(2, Labels).Value = "Greatest % Increase"
            .Cells(3, Labels).Value = "Greatest % Decrease"
            .Cells(4, Labels).Value = "Greatest Total Volume"
            .Cells(1, (Labels + 1)).Value = "Ticker"
            .Cells(1, (Labels + 2)).Value = "Value"
            'Inputting the greatest values
            .Cells(2, (Labels + 1)).Value = GreatestPerInName
            .Cells(2, (Labels + 2)).Value = GreatestPerIn
            .Cells(3, (Labels + 1)).Value = GreatestPerDecName
            .Cells(3, (Labels + 2)).Value = GreatestPerDec
            .Cells(4, (Labels + 1)).Value = GreatestVolName
            .Cells(4, (Labels + 2)).Value = GreatestVol
            'Formatting
            .UsedRange.EntireColumn.AutoFit
            .UsedRange.HorizontalAlignment = xlCenter
            .Columns((LastColumn + 4)).NumberFormat = "0.00%"
            .Cells(2, (Labels + 2)).NumberFormat = "0.00%"
            .Cells(3, (Labels + 2)).NumberFormat = "0.00%"
        End With
    Next ws
End Sub
