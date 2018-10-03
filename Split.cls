Sub Split():

' un merge

Dim cell As Range, joinedCells As Range

For Each cell In Range("E4:I60")
    If cell.MergeCells Then
        Set joinedCells = cell.MergeArea
        cell.MergeCells = False
        joinedCells.Value = cell.Value
    End If
Next


    Const NameCol = "B"
    Const HeaderRow = 3
    Const FirstRow = 4
    Dim SrcSheet As Worksheet
    Dim TrgSheet As Worksheet
    Dim SrcRow As Long
    Dim LastRow As Long
    Dim TrgRow As Long
    Dim Student As String
    Application.ScreenUpdating = False
    Set SrcSheet = ActiveSheet
    LastRow = SrcSheet.Cells(SrcSheet.Rows.Count, NameCol).End(xlUp).Row
    For SrcRow = FirstRow To LastRow
        Student = SrcSheet.Cells(SrcRow, NameCol).Value
        Set TrgSheet = Nothing
        On Error Resume Next
        Set TrgSheet = Worksheets(Student)
        On Error GoTo 0
        If TrgSheet Is Nothing Then
            Set TrgSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
            TrgSheet.Name = Student
            SrcSheet.Rows(HeaderRow).Copy Destination:=TrgSheet.Rows(HeaderRow)
        End If
        TrgRow = TrgSheet.Cells(TrgSheet.Rows.Count, NameCol).End(xlUp).Row + 1
        SrcSheet.Rows(SrcRow).Copy Destination:=TrgSheet.Rows(TrgRow)
    Next SrcRow
    Application.ScreenUpdating = True

End Sub
