Sub UnMergeFill()

Dim cell As Range, joinedCells As Range

For Each cell In Range("E4:I60")
    If cell.MergeCells Then
        Set joinedCells = cell.MergeArea
        cell.MergeCells = False
        joinedCells.Value = cell.Value
    End If
Next
End Sub
