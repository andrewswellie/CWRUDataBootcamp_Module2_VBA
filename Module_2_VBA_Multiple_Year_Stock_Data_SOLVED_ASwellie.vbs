Sub RunFinancials()

' Cycle through all the worksheets
For Each ws In Worksheets

' Declare and Assign variables
Dim Ticker As String
Dim volume As Double
Dim YearOpen As Double
Dim YearClose As Double
Dim LastRow As Long
Dim Summary As Integer
Summary = 2
volume = 0
YearOpen = ws.Cells(2, 3)
YearClose = 0


' Calculating the last Row on each spreadsheet
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Row and Column Formatting
ws.Range("I1") = "Ticker Symbol"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "% Change"
ws.Range("L1") = "Total Annual Volume"
ws.Range("J:J").NumberFormat = "$#,##0.00"
ws.Range("K:K").NumberFormat = "0.00%"
ws.Range("L:L").NumberFormat = "$#,##0.00"
ws.Columns("I:L").Columns.AutoFit


' For Loop to cycle through all the financial data and summarize volume,
' identify the initial open value, and identify the final close value of each symbol
For i = 2 To LastRow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
Ticker = ws.Cells(i, 1).Value
volume = volume + ws.Cells(i, 7).Value
ws.Range("I" & Summary).Value = Ticker
ws.Range("L" & Summary).Value = volume
ws.Range("J" & Summary).Value = YearClose - YearOpen

If YearOpen = 0 Then
ws.Range("K" & Summary).Value = "N/A"
Else: ws.Range("K" & Summary).Value = (YearClose - YearOpen) / YearOpen
End If

Summary = Summary + 1
volume = 0
YearOpen = ws.Cells(i + 1, 3).Value
YearClose = 0
Else
volume = volume + ws.Cells(i, 7).Value
YearClose = ws.Cells(i + 1, 6).Value
End If

Next i

Next ws

' calls next sub routine
Call FormatFinancials

End Sub


Sub FormatFinancials()

' Establish a new for loop to run through all sheets and determine last row in the new data set
For Each ws In Worksheets
LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row

' Set loop to add conditional formatting to othe % Change column
For i = 2 To LastRow
If ws.Range("J" & i) > 0 Then
ws.Range("J" & i).Interior.ColorIndex = 4
ElseIf ws.Range("J" & i) < 0 Then
ws.Range("J" & i).Interior.ColorIndex = 3
Else
ws.Range("J" & i).Interior.ColorIndex = 0
End If
            
Next i
Next ws

' calls next sub routine
Call FindMax

End Sub


Sub FindMax()

' Establish a new for loop to run through all sheets and determine last row in the new data set
For Each ws In Worksheets
LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row

' Row and Column formatting
ws.Range("N2") = "Greatest % Increase"
ws.Range("N3") = "Greatest % Decrease"
ws.Range("N4") = "Greatest Total Volume"
ws.Range("O1") = "Ticker"
ws.Range("P1") = "Value"

ws.Range("P2").NumberFormat = "0.00%"
ws.Range("P3").NumberFormat = "0.00%"
ws.Range("P4").NumberFormat = "$#,##0.00"

'Declare and assign new variables
Dim ChangeMax As Double
Dim ChangeMin As Double
Dim volume As Long


' Find the max of % increase and volume and the min for % increase
ws.Range("P2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & LastRow)) * 100
ws.Range("P3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & LastRow)) * 100
ws.Range("P4") = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))

' Determines the value of the min, max, and volume and assigns the value to a variable
ChangeMax = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & LastRow)), ws.Range("K2:K" & LastRow), 0)
ChangeMin = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & LastRow)), ws.Range("K2:K" & LastRow), 0)
volume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & LastRow)), ws.Range("L2:L" & LastRow), 0)

' Adds the ticker symbol for each item
ws.Range("O2") = ws.Cells(ChangeMax + 1, 9)
ws.Range("O3") = ws.Cells(ChangeMin + 1, 9)
ws.Range("O4") = ws.Cells(volume + 1, 9)

' formats data cells after values have been assigned
ws.Columns("N:P").Columns.AutoFit

Next ws

End Sub
