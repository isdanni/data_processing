Sub LoopAllExcelFiles()
'PURPOSE: To loop through all Excel files in a user specified folder and perform a set task on them
'SOURCE: www.TheSpreadsheetGuru.com

Dim wb As Workbook
Dim myPath As String
Dim myFile As String
Dim myExtension As String
Dim FldrPicker As FileDialog

'Optimize Macro Speed
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.Calculation = xlCalculationManual

'Retrieve Target Folder Path From User
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FldrPicker
      .Title = "Select A Target Folder"
      .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        myPath = .SelectedItems(1) & "\"
    End With

'In Case of Cancel
NextCode:
  myPath = myPath
  If myPath = "" Then GoTo ResetSettings

'Target File Extension (must include wildcard "*")
  myExtension = "*.xls*"

'Target Path with Ending Extention
  myFile = Dir(myPath & myExtension)

'Loop through each Excel file in folder
  Do While myFile <> ""
    'Set variable equal to opened workbook
      Set wb = Workbooks.Open(Filename:=myPath & myFile)
    
    'Ensure Workbook has opened before moving on to next line of code
      DoEvents
    
    '---------------------MY OPERATION-----------------------------
    
      ' 1. Un-Merge -----------------------------------------------
    Option Explicit
    Dim MainWorkBook As Workbook
    Dim NewWorkBook As Workbook
    Dim Pointer As Long
    Dim StrFile As String
    Dim cell As Range, joinedCells As Range

    For Each cell In Range("E4:I60")
        If cell.MergeCells Then
            Set joinedCells = cell.MergeArea
            cell.MergeCells = False
            joinedCells.Value = cell.Value
        End If
    Next
    
    
    ' 2. Split to Worksheets-------------------------------------------

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
    
    
    '--------------------------------FINISHED---------------------------
    
    'Save and Close Workbook
      wb.Close SaveChanges:=True
      
    'Ensure Workbook has closed before moving on to next line of code
      DoEvents

    'Get next file name
      myFile = Dir
  Loop

'Message Box when tasks are completed
  MsgBox "Task Complete!"

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub
