Function GetStringWithSubstringInArray(SubString As String, SourceArray As Variant, Optional CaseSensitive As Boolean = False) As String

    Dim i As Long
    
    For i = 0 To UBound(SourceArray)
        If VarType(SourceArray(i)) = vbString Then
            If StringContains(SourceString, SubString, CaseSensitive) Then
                GetStringWithSubstringInArray = SourceArray(i)
                Exit Function
            End If
        End If
    Next i

End Function
