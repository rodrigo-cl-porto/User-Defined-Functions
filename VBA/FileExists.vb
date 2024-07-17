Public Function FileExists(FilePath As String) As Boolean

    'PURPOSE: Test to see if a file exists or not
    'SOURCE: www.TheSpreadsheetGuru.com/The-Code-Vault
    'RESOURCE: http://www.rondebruin.nl/win/s9/win003.htm

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
