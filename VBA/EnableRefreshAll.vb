Public Sub EnableRefreshAll(ByRef wb As Workbook)
   
    Dim i As Long
    
    With wb
        For i = 1 To .Connections.Count
          If .Connections(i).Type = xlConnectionTypeOLEDB Then
            .Connections(i).RefreshWithRefreshAll = True
          End If
        Next i
    End With
    
End Sub
