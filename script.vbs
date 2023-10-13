Sub stock_script()
    'Loop through sheets
    For Each ws In ThisWorkbook.Worksheets
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ' Define Variables
        Dim Ticker As String
        Dim StockOpen As Double
        Dim StockClose As Double
        Dim Change As Double
        Dim TotalStockVolume As Double
        ' Create summary table
        Dim SummaryTableRow As Long
        SummaryTableRow = 2
        ' Set variables to zero
        PercentChange = 0
        TotalStockVolume = 0
        ' Loop through rows
        For I = 2 To LastRow
        ' Check if next row is same as current row. If not:
            If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
                ' Set ticker
                Ticker = Cells(I, 1).Value
                ' Set close
                StockClose = Cells(I, 6).Value
                ' Add to TotalStockVolume
                TotalStockVolume = TotalStockVolume + Cells(I, 7)
                ' Calculate Change
                Change = StockClose - StockOpen
                ' Print ticker in summary table
                Range("I" & SummaryTableRow).Value = Ticker
                ' Print Change to summary table
                Range("J" & SummaryTableRow).Value = Change
                ' Print PercentChange to summary table
                Range("K" & SummaryTableRow).Value = Change / StockOpen
                ' Print TotalStockVolume to summary table
                Range("L" & SummaryTableRow).Value = TotalStockVolume
                ' Add one to summary table row
                SummaryTableRow = SummaryTableRow + 1
                ' Reset Total Stock Volume
                TotalStockVolume = 0
            ' If first row with ticker
            ElseIf Cells(I - 1, 1).Value <> Cells(I, 1).Value Then
                ' Set open
                StockOpen = Cells(I, 3).Value
            ' If following and precedings rows are same ticker
            Else
                ' Add to TotalStockVolume
                TotalStockVolume = TotalStockVolume + Cells(I, 7)
            End If
        Next I
    Next ws
End Sub
