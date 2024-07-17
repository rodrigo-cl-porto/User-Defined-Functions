Public Function PreviousMonthNumber(dt As Date) As Integer

    Dim MonthNumber As Integer: MonthNumber = Month(dt)
    
    If MonthNumber = 1 Then
        PreviousMonthNumber = 12
    Else
        PreviousMonthNumber = MonthNumber - 1
    End If

End Function
