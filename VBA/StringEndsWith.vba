Function StringEndsWith(str1 As String, str2 As String, Optional caseSensitive As Boolean = False) As Boolean
    
    Dim Result As Boolean

    If caseSensitive Then
    
        If InStr(str1, str2) = Len(str1) - Len(str2) + 1 Then
            Result = True
        Else
            Result = False
        End If
        
    Else
    
        If InStr(1, str1, str2, vbTextCompare) = Len(str1) - Len(str2) + 1 Then
            Result = True
        Else
            Result = False
        End If
        
    End If

    StringEndsWith = Result

End Function
