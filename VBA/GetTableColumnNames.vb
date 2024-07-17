Public Function GetTableColumnNames(lo As ListObject) As String()

    Dim ColNames()   As String
    Dim i            As Long
    Dim TotalColumns As Long
    
    TotalColumns = lo.ListColumns.Count
    ReDim ColNames(TotalColumns - 1) As String
    
    
    For i = 0 To TotalColumns - 1
        ColNames(i) = lo.HeaderRowRange.Cells(i + 1)
    Next i
    
    GetTableColumnNames = ColNames

End Function
