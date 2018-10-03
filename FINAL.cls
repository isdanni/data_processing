Option Explicit
Dim MainWorkBook As Workbook
Dim NewWorkBook As Workbook
Dim Pointer As Long
Sub Split():

' Un-Merge

Dim cell As Range, joinedCells As Range

For Each cell In Range("E4:I60")
    If cell.MergeCells Then
        Set joinedCells = cell.MergeArea
        cell.MergeCells = False
        joinedCells.Value = cell.Value
    End If
Next

' Split to Worksheets

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
    

    Set MainWorkBook = ActiveWorkbook
    Range("E4").Value = MainWorkBook.Sheets.Count

    Application.ScreenUpdating = False   'enhance the performance
For Pointer = 2 To MainWorkBook.Sheets.Count
    Set NewWorkBook = Workbooks.Add
    MainWorkBook.Sheets(Pointer).Copy After:=NewWorkBook.Sheets(1)
    Application.DisplayAlerts = False
    NewWorkBook.Sheets(1).Delete
    Application.DisplayAlerts = True
    With NewWorkBook
        .SaveAs Filename:="D:\LIDA7\OneDrive - Orient Overseas Container Line Ltd\Projects\9. Hardware_List\data\VESSEL IT - Inventory\" & MainWorkBook.Sheets(Pointer).Name & ".xls"
    End With
    NewWorkBook.Close SaveChanges:=True
Next Pointer

Application.ScreenUpdating = True

End Sub

