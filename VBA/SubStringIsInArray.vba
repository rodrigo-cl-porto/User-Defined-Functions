Function SubstringIsInArray(subStr As String, srcArray As Variant, Optional caseSensitive As Boolean = False) As Boolean

    Dim i As Long
    
    For i = 0 To UBound(srcArray)
        If VarType(srcArray(i)) = vbString Then
            If StringContains(srcArray(i), subStr, caseSensitive) Then
                SubstringIsInArray = True
                Exit Function
            End If
        End If
    Next i

End Function
