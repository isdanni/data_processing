Public Sub PerformCopy()
    CopyFiles "source foler path\", "target folder path\"
End Sub

' First Macro
' Copy all types of files in subfolders to target folder
' Call CleanFile Macro next
' TESTING FINSIHED

Public Sub CopyFiles(ByVal strPath As String, ByVal strTarget As String)
Dim FSO As Object
Dim FileInFromFolder As Object
Dim FolderInFromFolder As Object
Dim Fdate As Long
Dim intSubFolderStartPos As Long
Dim strFolderName As String

Set FSO = CreateObject("scripting.filesystemobject")
    For Each FolderInFromFolder In FSO.GetFolder(strPath).SubFolders
            For Each FileInFromFolder In FSO.GetFolder(FolderInFromFolder).Files
                    FileInFromFolder.Copy strTarget
            Next
    Next

Call CleanFile
End Sub
