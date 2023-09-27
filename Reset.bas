Attribute VB_Name = "Reset"
Sub Reset()

Dim LastRow As Long
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 1 To LastRow

Cells(i, 9).Value = ""
Cells(i, 10).Value = ""
Cells(i, 11).Value = ""
Cells(i, 12).Value = ""

Next i

Range("J:J").Interior.ColorIndex = 2
Range("N1:P4").Value = ""

End Sub
