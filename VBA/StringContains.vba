Function StringContains(String1 As String, String2 As String, Optional CaseSensitive As Boolean = False) As Boolean
    
    If CaseSensitive Then
    
        If InStr(String1, String2) > 0 Then
            StringContains = True
        Else
            StringContains = False
        End If
        
    Else
    
        If InStr(1, String1, String2, vbTextCompare) > 0 Then
            StringContains = True
        Else
            StringContains = False
        End If
        
    End If

End Function
