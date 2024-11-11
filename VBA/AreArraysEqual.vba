Public Function AreArraysEqual(Array1 As Variant, Array2 As Variant) As Boolean

    Dim i As Long
    
    AreArraysEqual = True
    
    If UBound(Array1) = UBound(Array2) And LBound(Array1) = LBound(Array2) Then
        For i = 1 To UBound(Array1)
            If Array1(i) <> Array2(i) Then
                AreArraysEqual = False
                Exit For
            End If
        Next i
    Else
        AreArraysEqual = False
    End If

End Function
