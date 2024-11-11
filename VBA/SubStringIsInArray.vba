Function SubStringIsInArray(SubString As String, SourceArray As Variant, Optional CaseSensitive As Boolean = False) As Boolean

    Dim i As Long
    
    For i = 0 To UBound(SourceArray)
        
        If VarType(SourceArray(i)) = vbString Then
            If StringContains(SourceArray(i), SubString, CaseSensitive) Then
                SubStringIsInArray = True
                Exit Function
            End If
        End If
    
    Next i

End Function
