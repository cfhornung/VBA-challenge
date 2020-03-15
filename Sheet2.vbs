Sub Homework2()

'CFH
' This detects ticker changes then performs calculations

'turn off screen rewrite to improve performance
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'define variables
Dim Rownum As Long
Dim Blockstart As Long
Dim Reportrow As Long
Dim WS_Count As Integer
Dim Sheet As Integer

WS_Count = ActiveWorkbook.Worksheets.Count
'MsgBox WS_Count

'worksheet loop and selection
For Sheet = 1 To WS_Count
Worksheets(Sheet).Activate

'initialize, starting with row 2 to skip text headers
Blockstart = 2
Reportrow = 2
Rownum = 2

' write headers
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'find the maximum number of rows (for each sheet)
Dim Total As Long
Total = Cells(Rows.Count, 1).End(xlUp).Row
'MsgBox Total

' this For loop detects ticker changes in column A
For Rownum = 2 To Total

' compare the cell to the next cell
Range("A" & Rownum).Select
If ActiveCell.Value <> ActiveCell.Offset(1, 0).Value Then

' ticker (col I)
Range("A" & Rownum).Copy Range("I" & Reportrow)
'Range("H" & Rownum).Interior.ColorIndex = 19

' yearly change (col J)
Columns("J").ColumnWidth = 12
If Range("C" & Blockstart) <> 0 Then
Range("J" & Reportrow).Formula = "=F" & Rownum & " -C" & Blockstart & ""
End If

' percent change (col K)
Columns("K").ColumnWidth = 12
Range("K" & Reportrow).NumberFormat = "0.00%"

'Prevent divide by zero denominator condition
If Range("C" & Blockstart) <> 0 Then
Range("K" & Reportrow).Formula = "=(F" & Rownum & " -C" & Blockstart & ")/C" & Blockstart & ""
End If

' color percent change
If Range("K" & Reportrow).Value > 0 Then
Range("K" & Reportrow).Interior.ColorIndex = 4
End If

If Range("K" & Reportrow).Value < 0 Then
Range("K" & Reportrow).Interior.ColorIndex = 3
End If

' total volume (col L)
Columns("L").ColumnWidth = 15
Range("L" & Reportrow).NumberFormat = "0"
Range("L" & Reportrow).Formula = "=SUM(G" & Rownum & " :C" & Blockstart & ")"


'Re-initialize Blockstart to be the start of the next ticker block
Blockstart = Rownum + 1
Reportrow = Reportrow + 1

End If
Next Rownum

' extra credit summary
' write headers
Columns("M").ColumnWidth = 5
Columns("N").ColumnWidth = 5
Columns("O").ColumnWidth = 18
Columns("L").ColumnWidth = 20
Columns("Q").ColumnWidth = 20
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("Q2").NumberFormat = "0.00%"
Range("Q3").NumberFormat = "0.00%"

' find max and min
Dim Tickers
Tickers = Application.Range("I2", Range("I2").End(xlDown))

Dim Rng
Rng = Application.Range("K2", Range("K2").End(xlDown))

Dim Rng2
Rng2 = Application.Range("L2", Range("L2").End(xlDown))

Range("Q2").Value = Application.WorksheetFunction.Max(Rng)
Range("Q3").Value = Application.WorksheetFunction.Min(Rng)
Range("Q4").Value = Application.WorksheetFunction.Max(Rng2)

'now find the corresponding ticker to the max's and min's
Range("P2").Value = "???"
Range("P3").Value = "???"
Range("P4").Value = "???"

' next worksheet
 Next Sheet

'turn screen rerite back on
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

