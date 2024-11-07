Function Summation(Expression As String, First As Long, Last As Long) As Double

    'Evaluates a summation of a math expression
    'Summation("2*n-1", 1, 10) = 100
    'Summation("1/x^2", 1, 1000000) = 1,64493306684877 ≈ pi²/6

    Dim Rows As Long
    Dim Var  As String
    
    Var = Right(GetLettersOnly(Expression), 1)
    Rows = Last - First + 1
    
    Summation = Evaluate("=SUM(LET(" & Var & ", SEQUENCE(" & Rows & ", 1, " & First & "), " & Expression & "))")

End Function
