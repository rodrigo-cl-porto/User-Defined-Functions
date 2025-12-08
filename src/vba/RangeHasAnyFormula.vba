Public Function RangeHasAnyFormula(ByVal rng As Range) As Boolean

    Dim ErrNumber As Integer
    Dim ErrText   As String
    Dim Return    As Boolean

    If rng Is Nothing Then Exit Function

    On Error Resume Next
    rng.SpecialCells xlCellTypeFormulas
    ErrNumber = Err.Number
    ErrText = Err.Description
    On Error GoTo 0
    
    If ErrNumber = 0 Then
        Return = True
    ElseIf ErrText = "No cells were found." Then
        Return = False
    Else
        MsgBox "The following error occured: " & ErrNumber & vbLf & vbLf & ErrText, vbCritical + vbOKOnly, "Message Error"
    End If

    RangeHasAnyFormula = Return

End Function
