Function StringEndsWith(String1 As String, String2 As String, Optional CaseSensitive As Boolean = False) As Boolean
    
    If CaseSensitive Then
    
        If InStr(String1, String2) = Len(String1) - Len(String2) + 1 Then
            StringEndsWith = True
        Else
            StringEndsWith = False
        End If
        
    Else
    
        If InStr(1, String1, String2, vbTextCompare) = Len(String1) - Len(String2) + 1 Then
            StringEndsWith = True
        Else
            StringEndsWith = False
        End If
        
    End If

End Function
