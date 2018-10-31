
Public Sub PerformCopy()
    CopyFiles "D:\LIDA7\OneDrive - Orient Overseas Container Line Ltd\Projects\9. Hardware_List\TEST\VESSEL IT - Inventory\", "D:\LIDA7\OneDrive - Orient Overseas Container Line Ltd\Projects\9. Hardware_List\TEST\VESSEL IT - Inventory\TEST\"
End Sub


Public Sub CopyFiles(ByVal strPath As String, ByVal strTarget As String)
Dim FSO As Object
Dim FileInFromFolder As Object
Dim FolderInFromFolder As Object
Dim Fdate As Long
Dim intSubFolderStartPos As Long
Dim strFolderName As String

Set FSO = CreateObject("scripting.filesystemobject")
'First loop through files
    'For Each FileInFromFolder In FSO.GetFolder(strPath).Files
     '   Fdate = Int(FileInFromFolder.DateLastModified)
      '      FileInFromFolder.Copy strTarget
    'Next

    'Next loop throug folders
    For Each FolderInFromFolder In FSO.GetFolder(strPath).SubFolders
            For Each FileInFromFolder In FSO.GetFolder(FolderInFromFolder).Files
                    FileInFromFolder.Copy strTarget
            Next
    Next


Call ListAllFile
End Sub
