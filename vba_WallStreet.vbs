Sub vba_WallStreet()

'Set For All Worksheets
For Each ws In Worksheets

'Rearrange Order of Years
Worksheets("2014").Move _
before:=Worksheets("2015")
Worksheets("2016").Move _
after:=Worksheets("2015")

'Capitalise/Correct Headers
ws.Range("A1").Value = "Ticker"
ws.Range("B1").Value = "Date"
ws.Range("C1").Value = "Stock on Opening"
ws.Range("D1").Value = "Stock High"
ws.Range("E1").Value = "Stock Low"
ws.Range("F1").Value = "Stock on Close"
ws.Range("G1").Value = "Stock Volume"

'Set Column Headers Names
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'Autofit Column Widths
Worksheets("2014").Range("A:L").Columns.AutoFit
Worksheets("2015").Range("A:L").Columns.AutoFit
Worksheets("2016").Range("A:L").Columns.AutoFit

'Set Values
Dim TickerName As String
Dim YearlyOpen As Double
Dim YearlyClose As Double
Dim YearlyChange As Double
Dim TotalStockVolume As Double
TotalStockVolume = 0
Dim SummaryTableRow As Long
SummaryTableRow = 2
Dim LastRow As Long
Dim PriorResult As Long
PriorResult = 2

'Set Last Row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
For i = 2 To LastRow

'For Total Stock Volume
TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'Set Ticker Name
TickerName = ws.Cells(i, 1).Value

'Print The Ticker Name
ws.Range("I" & SummaryTableRow).Value = TickerName

'Print The Ticker Total Amount
ws.Range("L" & SummaryTableRow).Value = TotalStockVolume

'Reset Ticker Total
TotalStockVolume = 0

'Set Yearly Open
YearlyOpen = ws.Range("C" & PriorResult)

'Set Yearly Close
YearlyClose = ws.Range("F" & i)

'Set Yearly Change
YearlyChange = YearlyClose - YearlyOpen

ws.Range("J" & SummaryTableRow).Value = YearlyChange

'Determine Percent Change
If YearlyOpen = 0 Then
PercentChange = 0
Else
YearlyOpen = ws.Range("C" & PriorResult)
PercentChange = YearlyChange / YearlyOpen
End If

'Format Cells to Percentage
ws.Range("K" & SummaryTableRow).Value = PercentChange
ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"

'Highlight Cells- Positive(Green) or Negative(Red)
If ws.Range("J" & SummaryTableRow).Value >= 0 Then
ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
Else
ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
End If

'Repeat by Adding One To The Summary Table Row
SummaryTableRow = SummaryTableRow + 1
PriorResult = i + 1

End If

Next i

Next ws

End Sub
