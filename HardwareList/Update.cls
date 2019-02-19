Public Sub PerformCopy()
    ' Define source
    Dim xFileDialog As FileDialog
    Dim src As String
    On Error Resume Next
    Set xFileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    xFileDialog.AllowMultiSelect = False
    xFileDialog.Title = "Select a Source folder"
    src = xFileDialog.SelectedItems(1)
    
    ' Define dest
    Dim dest As String
    On Error Resume Next
    Set xFileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    xFileDialog.AllowMultiSelect = False
    xFileDialog.Title = "Select a Desttination folder"
    dest = xFileDialog.SelectedItems(1)
    
    ' =========================== COPY FILE ====================================
    CopyFiles src, dest

End Sub




Public Sub CopyFiles(src, dest)
    Set FSO = CreateObject("scripting.filesystemobject")
    
    'First loop through files
    For Each FileInFromFolder In FSO.getfolder(src).Files
    Fdate = Int(FileInFromFolder.DateLastModified)
    If Fdate >= Date - 1 Then
        FileInFromFolder.Copy strTarget
    End If

    'Next loop throug folders
    For Each FolderInFromFolder In FSO.getfolder(dest).SubFolders
        CopyFiles FolderInFromFolder.Path, strTarget
    Next folder
End Sub

