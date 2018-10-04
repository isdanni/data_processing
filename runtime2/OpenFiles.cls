Sub OpenFiles()
    Dim xStrPath As String
    Dim xFileDialog As FileDialog
    Dim xFile As String
    On Error Resume Next
    Set xFileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    xFileDialog.AllowMultiSelect = False
    xFileDialog.Title = "Select a folder"
    If xFileDialog.Show = -1 Then
        xStrPath = xFileDialog.SelectedItems(1)
    End If
    If xStrPath = "" Then Exit Sub
    xFile = Dir(xStrPath & "\*.xls")
    
    Do While xFile <> ""
        Call SplitData
    Loop
End Sub
