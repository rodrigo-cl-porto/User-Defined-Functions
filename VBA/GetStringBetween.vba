Function GetStringBetween(str As String, startStr As String, endStr As String) As String

    Dim match   As String
    Dim matches As Object
    Dim re      As Object: Set re = CreateObject("vbscript.regexp")
    
    With re
        .pattern = startStr & ".*?" & endStr
        .IgnoreCase = True
        .Global = False
        
        Set matches = .Execute(str)
        
        If matches.Count > 0 Then
            match = matches(0).value
            match = Replace(match, startStr, "")
            match = Replace(match, endStr, "")
        Else
            match = ""
        End If
    End With
    
    GetStringBetween = match

End Function
