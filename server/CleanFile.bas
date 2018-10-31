Sub CleanFile()
     
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim ws As Worksheet
     
    Set objFSO = CreateObject("Scripting.FileSystemObject")
   
    Set objFolder = objFSO.GetFolder("D:\LIDA7\OneDrive - Orient Overseas Container Line Ltd\Projects\9. Hardware_List\TEST\VESSEL IT - Inventory\TEST\")
   
     'Loop through the Files collection
    For Each objFile In objFolder.Files
        If Not InStr(1, objFile.Name, "Computer") > 0 Then
            Kill objFile
        End If
    Next
     
     'Clean up!
    Set objFolder = Nothing
    Set objFile = Nothing
    Set objFSO = Nothing
     
End Sub

