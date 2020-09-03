Sub vba_Challenge()

'Set For All Worksheets
For Each ws In Worksheets

'Print Column Headers Names
ws.Range("O1") = " "
ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total Volume"
ws.Range("P1") = "Ticker"
ws.Range("Q1") = "Value"

'Make Titles Bold
Worksheets("2014").Range("A1:Q1").Font.Bold = True
Worksheets("2015").Range("A1:Q1").Font.Bold = True
Worksheets("2016").Range("A1:Q1").Font.Bold = True

'Autofit Column Widths
Worksheets("2014").Range("O:Q").Columns.AutoFit
Worksheets("2015").Range("O:Q").Columns.AutoFit
Worksheets("2016").Range("O:Q").Columns.AutoFit

'Set Dimensions
Dim max As Double
max = 0
Dim min As Double
min = 0
Dim max_total_vol As Double
max_total_vol = 0
Dim min_row_index As Integer
Dim max_row_index As Integer
Dim max_total_vol_index As Integer
Dim LastRow As Long

'Set Last Row
LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

For i = 2 To LastRow

'Determine Higher/Lower for Min/Max
If ws.Cells(i, 11) > max Then
max = ws.Cells(i, 11)
max_row_index = i
End If

If ws.Cells(i, 11) < min Then
min = ws.Cells(i, 11)
min_row_index = i
End If

'Continue if Higher For Max Total Volume
If ws.Cells(i, 12) > max_total_vol Then
max_total_vol = ws.Cells(i, 12)
max_total_vol_index = i
End If
Next i

'Print Values
ws.Range("P2") = ws.Cells(max_row_index, 9).Value
ws.Range("P3") = ws.Cells(min_row_index, 9).Value
ws.Range("P4") = ws.Cells(max_total_vol_index, 9).Value

ws.Range("Q2") = max
ws.Range("Q3") = min
ws.Range("Q4") = max_total_vol

ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").NumberFormat = "0.00%"
        
Next ws

End Sub
