Public Sub DisableRefreshAll()
   
    Dim i As Long

    With ThisWorkbook
        For i = 1 To .Connections.Count
          'Excludes PowerPivot and other connections
          If .Connections(i).Type = xlConnectionTypeOLEDB Then
            .Connections(i).RefreshWithRefreshAll = False
          End If
        Next i
    End With
    
End Sub
