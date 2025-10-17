Function WorksheetHasListObject(ws As Worksheet) As Boolean

    Dim result As Boolean

    If ws.ListObjects.Count >= 1 Then
        result = True
    Else
        result = False
    End If

    WorksheetHasListObject = result

End Function
