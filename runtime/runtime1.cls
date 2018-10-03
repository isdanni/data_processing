Sub Split()

Dim Fso As Object, objFolder As Object, objSubFolder As Object
Dim FromPath As String
Dim FileInFolder As Object

FromPath = "D:\LIDA7\OneDrive - Orient Overseas Container Line Ltd\Projects\9. Hardware_List\backup\VESSEL IT - Inventory\"
Set Fso = CreateObject("Scripting.filesystemobject")
Set objFolder = Fso.GetFolder(FromPath)

For Each objSubFolder In objFolder.SubFolders
    For Each FileInFolder In objSubFolder.Files

        If InStr(1, FileInFolder.Name, ".xls") Then
            ' UN MERGE
            Dim cell As Range, joinedCells As Range

            For Each cell In Range("E4:I60")
                If cell.MergeCells Then
                    Set joinedCells = cell.MergeArea
                    cell.MergeCells = False
                    joinedCells.Value = cell.Value
                End If
            Next


            ' Split to worksheets
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
            
        End If

    Next FileInFolder
Next objSubFolder

End Sub
