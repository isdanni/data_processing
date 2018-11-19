Sub OpenFiles()
    ' Third Macro
    ' Used to open each excel file for processing, close after finished
    ' TESTING FINISHED
    Dim xStrPath As String
    Dim xFile As Variant
    Dim xExtension As String
    Dim wb As Workbook
    
    xStrPath = "folder path"
    xFile = Dir(xStrPath & "\*.xls")


    ' looping with dir when dir is called in sub will break the code
    ' solution: use first loop to store the filename
    Dim myArray() As String
    ReDim myArray(0)

    While (xFile <> "")
        ReDim Preserve myArray(UBound(myArray) + 1)
        myArray(UBound(myArray)) = xFile
        xFile = Dir()
    Wend
    
    ' second loop, used store array to call sub
    Dim n As Integer
    For n = 1 To UBound(myArray)
        Set wb = Workbooks.Open(Filename:=xStrPath & "\" & myArray(n)) 'open file
        Call SplitData
        wb.Close SaveChanges:=False
    Next
    
     'Do While Len(xFile) > 0
     '   Set wb = Workbooks.Open(Filename:=xStrPath & "\" & xFile) 'open file
     '   Call SplitData
     '   wb.Close SaveChanges:=False 'close the file
     '   xFile = Dir 'Get next file name
     'Loop
    
End Sub
