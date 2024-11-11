Function StringStartsWith(String1 As String, String2 As String, Optional CaseSensitive As Boolean = False) As Boolean

    If CaseSensitive Then
    
        If InStr(String1, String2) = 1 Then
            StringStartsWith = True
        Else
            StringStartsWith = False
        End If
        
    Else
    
        If InStr(1, String1, String2, vbTextCompare) = 1 Then
            StringStartsWith = True
        Else
            StringStartsWith = False
        End If
        
    End If

End Function
