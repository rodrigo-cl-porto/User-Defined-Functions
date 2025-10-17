Function RangeHasConstantValues(rng As Range) As Boolean

    Dim ErrNumber As Integer
    Dim ErrText   As String
    Dim Result    As Boolean

    If rng Is Nothing Then Exit Function

    On Error Resume Next
    rng.SpecialCells (xlCellTypeConstants)
    ErrNumber = Err.Number
    ErrText = Err.Description
    On Error GoTo 0
    
    If ErrNumber = 0 Then
        Result = True
    ElseIf ErrText = "No cells were found." Then
        Result = False
    Else
        MsgBox "The following error occured: " & ErrNumber & vbLf & vbLf & ErrText, vbCritical + vbOKOnly, "Message Error"
    End If

    RangeHasConstantValues = Result

End Function
