Attribute VB_Name = "ForEachTicker"
Sub ForEachWS()
Dim ws As Worksheet

For Each ws In Worksheets
Call Ticker(ws)

Next ws

End Sub


'^Up above should be an independent command to repeat Ticker macro on each worksheet
'Alternatively there's a For Each loop below setting the variables and at the end
            ' Using msgBox and Next

'If adding ws. in front of objects doesn't make it work, remove sub start/stop for Ticker
' and put Next ws and End Sub at very bottom

'Before i put in the ws. in front of objects, Matthew T. said
'"Does ticker need () at the end when called? I can't remember"
'in regards to line 5

'------------------------------------------------




Sub Ticker(ws)
'Dim Current As Worksheet



'Define variables
Dim LastRow As Long
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Dim i As Long
Dim r As Long

Dim Ticker As String
Dim YearOpen As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim VolumeCounter As LongLong
VolumeCounter = 0
Dim SummaryRow As Integer
SummaryRow = 2
Dim GreatestIncrease As Double
GreatestIncrease = 0
Dim GreatestDecrease As Double
GreatestDecrease = 0
Dim GreatestVolume As LongLong
GreatestVolume = 0

'loop through worksheets
'For Each Current In Worksheets


'Write out column headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'Summary of summary rows/column identifiers
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest  % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"

'---------------------------------------------------------
'Data Manipulation
For i = 2 To LastRow
    'Store ticker
    Ticker = ws.Cells(i, 1).Value
    YearOpen = ws.Cells(i, 3).Value
    
    'Determine if ticker column is different
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    Ticker = ws.Cells(i, 1).Value
    'Store New Ticker and YearOpen
    YearOpen = ws.Cells(i, 3).Value
    VolumeCounter = 0
    'Increase SummaryRow by 1
    SummaryRow = SummaryRow + 1

Else
    YearOpen = ws.Cells(i, 3).Value
    YearlyChange = ws.Cells(i, 6).Value - YearOpen
    PercentChange = (YearlyChange / YearOpen)
    
    'Add volume to volume counter
    VolumeCounter = VolumeCounter + ws.Cells(i, 7).Value
'print summary to Right
    ws.Range("I" & SummaryRow).Value = Ticker
    ws.Range("J" & SummaryRow).Value = YearlyChange
    ws.Range("K" & SummaryRow).Value = PercentChange
    ws.Range("K" & SummaryRow).NumberFormat = "0.00%"
    ws.Range("L" & SummaryRow).Value = VolumeCounter
    End If

'Formatting YearlyChange
If ws.Range("J" & SummaryRow).Value > 0 Then
    ws.Range("J" & SummaryRow).Interior.ColorIndex = 4
    End If
If ws.Range("J" & SummaryRow).Value < 0 Then
    ws.Range("J" & SummaryRow).Interior.ColorIndex = 3
    End If
    
'---------------------------------------------------------
'looping through the summary data within iterations of i loop
'Returned GreatestIncrease and GreatestDecrease comparing any day of the year to the start
'rather than comparing the end of year to the start

'Summary of summary table (Greatest increase, decrease, volume)
'If Range("K" & SummaryRow).Value > GreatestIncrease Then
'    GreatestIncrease = Range("K" & SummaryRow).Value
'    Range("O2").Value = Ticker
'    Range("P2").Value = GreatestIncrease
'    Range("P2").NumberFormat = "0.00%"

'ElseIf Range("K" & SummaryRow).Value < GreatestDecrease Then
'    GreatestDecrease = Range("K" & SummaryRow).Value
'    Range("O3").Value = Ticker
'    Range("P3").Value = GreatestDecrease
'    Range("P3").NumberFormat = "0.00%"
'End If

'If Range("L" & SummaryRow).Value > GreatestVolume Then
'    GreatestVolume = Range("L" & SummaryRow).Value
'    Range("O4").Value = Ticker
'    Range("P4").Value = GreatestVolume
'End If
           
Next i

'---------------------------------------------------------
'My attempt at looping through the summary table after it's been made
'Lots of headaches, error messages, and occassionally returning null value ""
'to the summary of summary table

'Summary of summary table (Greatest increase, decrease, volume)
'For r = 2 To LastRow
'    Ticker = Cells(i, 8).Value
'    SummaryRow = 2
    
'If Range("K" & SummaryRow).Value > GreatestIncrease Then
'    GreatestIncrease = Range("K" & SummaryRow).Value
'    Range("O2").Value = Ticker
'    Range("P2").Value = GreatestIncrease
'    Range("P2").NumberFormat = "0.00%"

'ElseIf Range("K" & SummaryRow).Value < GreatestDecrease Then
'    GreatestDecrease = Range("K" & SummaryRow).Value
'    Range("O3").Value = Ticker
'    Range("P3").Value = GreatestDecrease
'    Range("P3").NumberFormat = "0.00%"
'End If

'If Range("L" & SummaryRow).Value > GreatestVolume Then
'    GreatestVolume = Range("L" & SummaryRow).Value
'    Range("O4").Value = Ticker
'    Range("P4").Value = GreatestVolume
'End If

'Increase SummaryRow by 1
'SummaryRow = SummaryRow + 1

'Next r


'---------------------------------------------------------
'From askBCS ask-279449. Sandile recommended using worksheet functions.
'I took the code she typed up and simplified it down.
'I got everything to work except returning the ticker for the GreatestVolume

'GreatestIncrease = WorksheetFunction.Max(Range("K2:K" & LastRow))
'    Range("P2").Value = GreatestIncrease
'    Range("P2").NumberFormat = "0.00%"
'    Range("O2").Value = WorksheetFunction.XLookup(GreatestIncrease, Range("K2:K91"), Range("I2:I91"))

'GreatestDecrease = WorksheetFunction.Min(Range("K2:K" & LastRow))
'    Range("O3").Value = WorksheetFunction.XLookup(GreatestDecrease, Range("K2:K91"), Range("I2:I91"))
'    Range("P3").Value = GreatestDecrease
'    Range("P3").NumberFormat = "0.00%"

'GreatestVolume = WorksheetFunction.Max(Range("L2:L" & LastRow))
'    Range("P4").Value = GreatestVolume
'    Range("O4").Value = WorksheetFunction.XLookup(GreatestVolume, Range("K2:K91"), Range("I2:I91"))
'^This row gives a RTE 1004 "Unable to get the xlookup property of the worksheet function class"


'---------------------------------------------------------
'Modified from https://raw.githubusercontent.com/qyubrooks/Multiple-Year-Stock-Data/master/VBA%20Script%20Hard
'FROM 'set initial variables... TO Range("P2:P3").NumberFormat = "0.00%" is from qyubrooks.
'Modifications include: removing any "ws.". I am looping through worksheets using another method
            'Add'l modifications- adding back in ws. in front of objects.
'Citation: GitHub. (n.d.). https://raw.githubusercontent.com/qyubrooks/Multiple-Year-Stock-Data/master/VBA%20Script%20Hard
      
      'set initial variables for holding greatest % increase, greatest % decrease, greatest total volume and their ticker names
      Dim Greatest_Increase As Double
      Dim Greatest_Decrease As Double
      Dim Greatest_Total As Double
      Dim Greatest_Increase_Ticker As String
      Dim Greatest_Decrease_Ticker As String
      Dim Greatest_Total_Ticker As String

        'set the first ticker's percent change and total stock volume as the greatest ones
        Greatest_Increase = ws.Range("K2").Value
        Greatest_Decrease = ws.Range("K2").Value
        Greatest_Total = ws.Range("L2").Value
        
        'Define last row of Ticker column
        Lastrow_Ticker = ws.Cells(Rows.Count, "I").End(xlUp).Row
        
        'Loop through each row of Ticker column to find the greatest results
         For r = 2 To Lastrow_Ticker:
               If ws.Range("K" & r + 1).Value > Greatest_Increase Then
                  Greatest_Increase = ws.Range("K" & r + 1).Value
                  Greatest_Increase_Ticker = ws.Range("I" & r + 1).Value
               ElseIf ws.Range("K" & r + 1).Value < Greatest_Decrease Then
                  Greatest_Decrease = ws.Range("K" & r + 1).Value
                  Greatest_Decrease_Ticker = ws.Range("I" & r + 1).Value
                ElseIf ws.Range("L" & r + 1).Value > Greatest_Total Then
                  Greatest_Total = ws.Range("L" & r + 1).Value
                  Greatest_Total_Ticker = ws.Range("I" & r + 1).Value
                End If
            Next r
            
            'Print greatest % increase, greatest % decrease, greatest total volume and their ticker names
            ws.Range("O2").Value = Greatest_Increase_Ticker
            ws.Range("O3").Value = Greatest_Decrease_Ticker
            ws.Range("O4").Value = Greatest_Total_Ticker
            ws.Range("P2").Value = Greatest_Increase
            ws.Range("P3").Value = Greatest_Decrease
            ws.Range("P4").Value = Greatest_Total
            ws.Range("P2:P3").NumberFormat = "0.00%"

'Tell user which sheet you're on
'MsgBox (Current)

'Move to the next worksheet
'Next Current

End Sub





