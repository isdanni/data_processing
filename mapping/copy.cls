Sub copy()
Dim ws1 As Worksheet
Dim ws2 As Worksheet
Dim ws3 As Worksheet
Dim r As Integer

Dim lRow1 As Long, lRow2 As Long, i As Long, j As Long
Set ws1 = ThisWorkbook.Sheets("nmapresult")
Set ws2 = ThisWorkbook.Sheets("list")
' Set ws3 = ThisWorkbook.Sheets("Sheet3")

lRow1 = ThisWorkbook.Sheets("list").Range("C" & Rows.Count).End(xlUp).Row
lRow2 = ThisWorkbook.Sheets("nmapresult").Range("B" & Rows.Count).End(xlUp).Row

r = 1

For i = 1 To lRow1
    For j = 1 To lRow2
        If ws2.Cells(i, 3).Value = ws1.Cells(j, 2).Value Then
            ws2.Cells(i, 1).copy Destination:=ws1.Cells(j, 4)
            ws2.Cells(i, 2).copy Destination:=ws1.Cells(j, 5)
            ws2.Cells(i, 4).copy Destination:=ws1.Cells(j, 6)
        r = r + 1

        End If
    Next j
Next i
End Sub
