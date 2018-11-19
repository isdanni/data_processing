 Sub SplitData()
' Fourth Macro
' Used to split worksheet data
' TESTING

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
            Dim Device As String
            Application.ScreenUpdating = False
            Set SrcSheet = ActiveSheet
            LastRow = SrcSheet.Cells(SrcSheet.Rows.Count, NameCol).End(xlUp).Row
            For SrcRow = FirstRow To LastRow
                If IsEmpty(SrcSheet.Cells(SrcRow, NameCol).Value) Then Exit For
                
                Device = SrcSheet.Cells(SrcRow, NameCol).Value
                Set TrgSheet = Nothing
                On Error Resume Next
                Set TrgSheet = Worksheets(Device)
                On Error GoTo 0
                If TrgSheet Is Nothing Then
                    Set TrgSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
                    TrgSheet.Name = Device
                    SrcSheet.Rows(HeaderRow).Copy Destination:=TrgSheet.Rows(HeaderRow)
                End If
            TrgRow = TrgSheet.Cells(TrgSheet.Rows.Count, NameCol).End(xlUp).Row + 1
            SrcSheet.Rows(SrcRow).Copy Destination:=TrgSheet.Rows(TrgRow)
            Next SrcRow
            ' NO SAVE!
            Application.ScreenUpdating = True
            
            
            ' Export worksheet
            Dim Pointer As Long
            Dim FilePath As String
            Dim Filename As String
            Set MainWorkBook = ActiveWorkbook
            Range("E4").Value = MainWorkBook.Sheets.Count

            Application.ScreenUpdating = False   'enhance the performance
            For Pointer = 2 To MainWorkBook.Sheets.Count
                Set NewWorkbook = Workbooks.Add
                MainWorkBook.Sheets(Pointer).Copy After:=NewWorkbook.Sheets(1)
                Application.DisplayAlerts = False
                NewWorkbook.Sheets(1).Delete
                Application.DisplayAlerts = False
                With NewWorkbook
                    Filename = "D:\LIDA7\OneDrive - Orient Overseas Container Line Ltd\Projects\9. Hardware_List\TEST\" & MainWorkBook.Sheets(Pointer).Name & ".xls"
                    FilePath = Dir(Filename)
                    
                    ' if file does not exist, save as new file name
                    If FilePath = "" Then
                        .SaveAs Filename
                        NewWorkbook.Close (0)
                    ' if file exists, copy the new workbook content to the existing file
                    Else
                        Dim newlast As String   ' new workbook last row
                        Dim originlast As String
                        Dim wb As Workbook
                        Dim rng1 As Range
                        
                        Set NewWorkbook = ActiveWorkbook
                        ' select the current new workbook data
                        newlast = NewWorkbook.Sheets(1).Cells(Sheets(1).Rows.Count, "B").End(xlUp).Row
                        Set rng1 = Range("A4:I" & newlast)
                        Range("A4:I" & newlast).Copy
                        'rng1.Select
                        'Selection.Copy
                        'NewWorkbook.Close (0)
                        
                        ' paste in existing file's last row
                        Set wb = Workbooks.Open(Filename)
                        originlast = wb.Sheets(1).Cells(Sheets(1).Rows.Count, "B").End(xlUp).Row
                        
                        wb.Sheets(1).Range("A" & originlast).Insert Shift:=xlDown
                        'wb.Sheets(1).Range("B" & originlast:"K1000").PasteSpecial xlPasteAll
                        Application.DisplayAlerts = False
                        wb.Close True
                        
                    End If
                End With
            Next Pointer
            Application.ScreenUpdating = True
End Sub

