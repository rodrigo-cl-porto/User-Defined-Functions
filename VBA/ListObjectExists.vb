Public Function ListObjectExists(ByRef wb As Workbook, ByVal loName As String) As Boolean
    
    Dim lo As ListObject
    Dim ws As Worksheet
    
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If lo.name = loName Then
                ListObjectExists = True
                Exit Function
            End If
        Next lo
    Next ws

End Function
