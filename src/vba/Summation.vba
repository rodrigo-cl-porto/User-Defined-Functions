Private Function GetLettersOnly(Text As String) As String

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

Public Function Summation(Expression As String, First As Long, Last As Long) As Double

    Dim Rows As Long
    Dim Var  As String
    
    Var = Right(GetLettersOnly(Expression), 1)
    Rows = Last - First + 1
    
    Summation = Evaluate("=SUM(LET(" & Var & ", SEQUENCE(" & Rows & ", 1, " & First & "), " & Expression & "))")

End Function
