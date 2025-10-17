Function Summation(Expression As String, First As Long, Last As Long) As Double

    Dim Rows As Long
    Dim Var  As String
    
    Var = Right(GetLettersOnly(Expression), 1)
    Rows = Last - First + 1
    
    Summation = Evaluate("=SUM(LET(" & Var & ", SEQUENCE(" & Rows & ", 1, " & First & "), " & Expression & "))")

End Function
