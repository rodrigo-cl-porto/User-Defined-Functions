Public Function RangeHasConstantValues(rng As Range) As Boolean

    Dim ErrNumber As Integer
    Dim ErrText   As String

    If rng Is Nothing Then
        RangeHasConstantValues = False
        Exit Function
    End If

    On Error Resume Next
    rng.SpecialCells (xlCellTypeConstants)
    ErrNumber = Err.Number
    ErrText = Err.Description
    On Error GoTo 0
    
    If ErrNumber = 0 Then
        RangeHasConstantValues = True
    ElseIf ErrText = "No cells were found." Then
        RangeHasConstantValues = False
    Else
        MsgBox "The following error occured: " & ErrNumber & vbLf & vbLf & ErrText, vbCritical + vbOKOnly, "Message Error"
    End If

End Function
