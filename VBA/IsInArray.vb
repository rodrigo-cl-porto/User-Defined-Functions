Public Function IsInArray(ValueToBeFound As Variant, SourceArray As Variant) As Boolean
    
    Dim i As Long
    
    For i = LBound(SourceArray) To UBound(SourceArray)
        If SourceArray(i) = ValueToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i

End Function
