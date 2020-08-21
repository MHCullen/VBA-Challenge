Sub Homework2()

'INSERT FOR EACH STUFF HERE
For Each ws In Worksheets

'Insert Column Header Data via Ranges
ws.Range("j1").Value = "Ticker"
ws.Range("k1").Value = "Yearly $ Change"
ws.Range("L1").Value = "Percent Change"
ws.Range("m1").Value = "Total Stock Volume"

'Set an initial variable for holding the Ticker Symbol
Dim Ticker As String

'Set an intital variable for holding the Total Volume
Dim Volume_Total As Double
Volume_Total = 0

'Keep track of the location for each Ticker Symbol in the Summary Table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2 'starting with second row

'Create another variable to help us first open and last price for percent change and $ change questions
Dim Start As Long
Start = 2

Dim Change As Double
Dim PercentChg As Double

'INSERT LAST ROW STUFF HERE
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'Loop through all the tickers with Last Row
For i = 2 To LastRow

'Check if we are on same ticker, if it is not....
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'Set the ticker
Ticker = ws.Cells(i, 1).Value

'Add to the Volume Total
Volume_Total = Volume_Total + ws.Cells(i, 7).Value 'Volume column is 7

'Print the Ticker in the Summary Table
'Original: Range("J" & Summary_Table_Row).Value = Ticker
ws.Cells(Summary_Table_Row, 10).Value = Ticker


'Print the Total Volume in the Summary Table
'Original: Range("M" & Summary_Table_Row).Value = Volume_Total
ws.Cells(Summary_Table_Row, 13).Value = Volume_Total

'Calculate $ Change (Last-First)
Change = ws.Cells(i, 6) - Cells(Start, 3)

'Calculate Percent Change
PercentChg = Change / ws.Cells(Start, 3)

Start = (i + 1)

'Print Change
ws.Cells(Summary_Table_Row, 11).Value = Change

'Print PercentChg
ws.Cells(Summary_Table_Row, 12).Value = PercentChg

'Conditional Formatting (remove if a problem)..taken from Student Gradebook exercise
If ws.Cells(Summary_Table_Row, 11).Value >= 0 Then
ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 4

Else
ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 3

End If

'Add one to the summary table row
Summary_Table_Row = Summary_Table_Row + 1

'Reset the Volume total
Volume_Total = 0

'If the cell immediately following a row is the same ticker
Else

'add to the Volume total
Volume_Total = Volume_Total + ws.Cells(i, 7).Value

End If

Next i

'Format Percentage
For i = 2 To LastRow
ws.Cells(i, 12).Style = "Percent"

Next i

Next ws


End Sub


