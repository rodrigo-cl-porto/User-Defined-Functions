Function RangeIsHidden(rng As Range) As Boolean

    Dim ErrNumber As Integer
    Dim ErrText   As String
    
    If rng Is Nothing Then
        RangeIsHidden = True
        Exit Function
    End If
    
    On Error Resume Next
    rng.SpecialCells (xlCellTypeVisible)
    ErrNumber = Err.Number
    ErrText = Err.Description
    On Error GoTo 0
    
    If ErrNumber = 0 Then
        RangeIsHidden = False
    ElseIf ErrText = "No cells were found." Then
        RangeIsHidden = True
    Else
        MsgBox "The following error occured: " & ErrNumber & vbLf & vbLf & ErrText, vbCritical + vbOKOnly, "Message Error"
    End If

End Function
