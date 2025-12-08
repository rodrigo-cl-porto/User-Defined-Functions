Public Function GetLettersOnly(Text As String) As String

    Dim i       As Long
    Dim Letters As String
    Dim Chr     As String

    For i = 1 To Len(Text)
    
        Chr = LCase(Mid(Text, i, 1))

        If Asc(Chr) >= 97 And Asc(Chr) <= 122 Then
            Letters = Letters + Chr
        End If
        
    Next i
    
    GetLettersOnly = Letters
    
End Function
