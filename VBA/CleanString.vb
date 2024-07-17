Public Function CleanString(ByVal myString As String, Optional ReplaceBySpace As Boolean = True, Optional ConvertNonBreakingSpace As Boolean = True) As String
    
    Dim i            As Long
    Dim CharsToClean As Variant
    
    CharsToClean = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 127, 129, 141, 143, 144, 157)
  
    If ConvertNonBreakingSpace Then myString = Replace(myString, Chr(160), " ")
    
    
    For i = LBound(CharsToClean) To UBound(CharsToClean)
        If InStr(myString, Chr(CharsToClean(i))) Then
            If ReplaceBySpace Then
                myString = Replace(myString, Chr(CharsToClean(i)), " ")
            Else
                myString = Replace(myString, Chr(CharsToClean(i)), "")
            End If
        End If
    Next
    
    CleanString = Trim(myString)
    
End Function
