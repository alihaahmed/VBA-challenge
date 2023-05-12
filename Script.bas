Attribute VB_Name = "Module1"
Sub module2challenge()

'Loop through worksheets

For Each ws In Worksheets

'Determine last row in each worksheet and store in "lastrow" variable

Dim lastrow As Long
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Create headers for new columns

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'Declare and set variables to be calculated

Dim tickername As String
tickername = " "
Dim yearlychange As Double
yearlychange = 0
Dim percentchange As Double
percentchange = 0
Dim volumetotal As Double
volumetotal = 0
Dim yearopenvalue As Double
yearopenvalue = 0
Dim yearclosevalue As Double
yearclosevalue = 0
Dim summarytablerow As Long
summarytablerow = 2

Dim maxpercentticker As String
Dim minpercentticker As String
Dim maxvolumeticker As String
Dim maxpercent As Double
Dim minpercent As Double
Dim maxvolume As Double

'Set initial opening value for first stock/ticker

yearopenvalue = ws.Cells(2, 3).Value

'Loop from second row to last row for each worksheet

For i = 2 To lastrow

'Check if on the same ticker; if not, store ticker name, calculate yearly change and percent change, and sum total stock volume

If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
tickername = ws.Cells(i, 1).Value
yearclosevalue = ws.Cells(i, 6).Value
yearlychange = yearclosevalue - yearopenvalue
percentchange = (yearlychange / yearopenvalue)
volumetotal = volumetotal + ws.Cells(i, 7).Value

'Print stored and calculated values in new columns

ws.Range("I" & summarytablerow).Value = tickername

'If yearly change is negative, format cell to red colour; if positive, format cell to green colour

ws.Range("J" & summarytablerow).Value = yearlychange
If yearlychange > 0 Then
ws.Range("J" & summarytablerow).Interior.ColorIndex = 4
ElseIf yearlychange < 0 Then
ws.Range("J" & summarytablerow).Interior.ColorIndex = 3
End If

'Format percent change as %

ws.Range("K" & summarytablerow).Value = percentchange
ws.Range("K" & summarytablerow).NumberFormat = "0.00%"

'Format total stock volume as whole number

ws.Range("L" & summarytablerow).Value = volumetotal
ws.Range("L" & summarytablerow).NumberFormat = "0"

'Reset values

summarytablerow = summarytablerow + 1
yearopenvalue = ws.Cells(i + 1, 3).Value
percentchange = 0
volumetotal = 0

'If still on same ticker, continue adding current volume to total stock volume

Else
volumetotal = volumetotal + ws.Cells(i, 7).Value

End If

Next i

'Label columns and rows

ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

'Calculate greatest % increase, print in associated cell for all worksheets with % format

maxpercent = ws.Application.WorksheetFunction.Max(Range("K:K"))
ws.Cells(2, 17).Value = maxpercent
ws.Cells(2, 17).NumberFormat = "0.00%"

'Calculate greatest % decrease, print in associated cell for all worksheets with % format

minpercent = ws.Application.WorksheetFunction.Min(Range("K:K"))
ws.Cells(3, 17).Value = minpercent
ws.Cells(3, 17).NumberFormat = "0.00%"

'Calculate greatest total volume, print in associated cell for all worksheets

maxvolume = ws.Application.WorksheetFunction.Max(Range("l:l"))
ws.Cells(4, 17).Value = maxvolume

'Determine tickers associated with calculated variables and print in associated cells

For i = 2 To lastrow

If ws.Cells(i, 11).Value = maxpercent Then
maxpercentticker = ws.Cells(i, 9).Value
ws.Cells(2, 16).Value = maxpercentticker
End If

If ws.Cells(i, 11).Value = minpercent Then
minpercentticker = ws.Cells(i, 9).Value
ws.Cells(3, 16).Value = minpercentticker
End If

If ws.Cells(i, 12).Value = maxvolume Then
maxvolumeticker = ws.Cells(i, 9).Value
ws.Cells(4, 16).Value = maxvolumeticker
End If

Next i

Next ws

End Sub
