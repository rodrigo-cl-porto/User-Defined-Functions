Function WorksheetHasListObject(ws As Worksheet) As Boolean

    If ws.ListObjects.Count >= 1 Then
        WorksheetHasListObject = True
    Else
        WorksheetHasListObject = False
    End If
    
End Function
