Public Function AreArraysEqual(Array1 As Variant, Array2 As Variant) As Boolean

    Dim i As Long
    Dim Return as Boolean
    
    Return = True
    
    If UBound(Array1) = UBound(Array2) And LBound(Array1) = LBound(Array2) Then
        For i = LBound(Array1) To UBound(Array1)
            If Array1(i) <> Array2(i) Then
                Return = False
                Exit For
            End If
        Next i
    Else
        Return = False
    End If

    AreArraysEqual = Return

End Function
