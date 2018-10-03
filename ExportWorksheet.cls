Option Explicit
Dim MainWorkBook As Workbook
Dim NewWorkBook As Workbook

Sub ExportWorksheet()
Dim Pointer As Long

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
