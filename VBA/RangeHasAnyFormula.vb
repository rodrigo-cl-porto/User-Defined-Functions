Public Function RangeHasAnyFormula(ByVal rng As Range) As Boolean

    Dim ErrNumber As Integer
    Dim ErrText   As String

    If rng Is Nothing Then
        RangeHasAnyFormula = False
        Exit Function
    End If

    On Error Resume Next
    rng.SpecialCells xlCellTypeFormulas
    ErrNumber = Err.Number
    ErrText = Err.Description
    On Error GoTo 0
    
    If ErrNumber = 0 Then
        RangeHasAnyFormula = True
    ElseIf ErrText = "No cells were found." Then
        RangeHasAnyFormula = False
    Else
        MsgBox "The following error occured: " & ErrNumber & vbLf & vbLf & ErrText, vbCritical + vbOKOnly, "Message Error"
    End If

End Function
