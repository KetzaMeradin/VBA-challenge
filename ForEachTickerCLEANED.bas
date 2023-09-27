Attribute VB_Name = "ForEachTicker"
Sub ForEachWS()
Dim ws As Worksheet

For Each ws In Worksheets
Call Ticker(ws)

Next ws

End Sub

Sub Ticker(ws)

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
               
Next i

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

End Sub





