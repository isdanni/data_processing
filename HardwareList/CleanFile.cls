Sub CleanFile()
    ' Second Macro
    ' Used to clean up unrelative files copied to the target folder
    ' TESTING FINISHED
     
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim ws As Worksheet
     
    Set objFSO = CreateObject("Scripting.FileSystemObject")
   
    Set objFolder = objFSO.GetFolder("folder path\")
   
     'Loop through the Files collection
     ' if file does not contain "Computer"
     ' clean up
    For Each objFile In objFolder.Files
        If Not InStr(1, objFile.Name, "Computer") > 0 Then
            Kill objFile
        End If
    Next
     
     'Clean up!
    Set objFolder = Nothing
    Set objFile = Nothing
    Set objFSO = Nothing
    
    Call OpenFiles
     
End Sub
