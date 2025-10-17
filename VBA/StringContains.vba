Function StringContains(str1 As String, str2 As String, Optional caseSensitive As Boolean = False) As Boolean
    
    Dim Return as String

    If caseSensitive Then
    
        If InStr(str1, str2) > 0 Then
            Return = True
        Else
            Return = False
        End If
        
    Else
    
        If InStr(1, str1, str2, vbTextCompare) > 0 Then
            Return = True
        Else
            Return = False
        End If
        
    End If

    StringContains = Return

End Function
