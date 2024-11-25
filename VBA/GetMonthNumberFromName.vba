Public Function GetMonthNumberFromName(MonthName As String) As Integer
    
    GetMonthNumberFromName = Application.Evaluate("=MONTH(1&""" & MonthName & """)")

End Function
