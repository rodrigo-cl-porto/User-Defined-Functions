Public Function TableHasQuery(tbl As ListObject) As Boolean

    If tbl Is Nothing Then
       TableHasQuery = False
       Exit Function
    End If

    On Error GoTo ErrHandler
    If Not (tbl.QueryTable Is Nothing) Then
        TableHasQuery = True
    End If
    
    Exit Function

ErrHandler:
    If Err.Number = 1004 Then 'Application-defined or object-defined error
        TableHasQuery = False
        On Error GoTo 0
    Else
        On Error GoTo 0
    End If
    
End Function
