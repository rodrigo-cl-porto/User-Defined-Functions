Public Function FileExists(FilePath As String) As Boolean

    Dim TestStr As String

    'Test File Path
    TestStr = Dir(FilePath)

    'Determine if File exists
    If TestStr = vbNullString Then
        FileExists = False
    Else
        FileExists = True
    End If

End Function
