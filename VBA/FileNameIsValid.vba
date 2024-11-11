Public Function FileNameIsValid(FileName As String) As Boolean

    'PURPOSE: Determine If A Given File Name Is Valid
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
    'AUTHOR: Jon Peltier

    Const strBadChar As String = "\/:*?<>|[]"""
    Dim i            As Long

    'Assume valid unless it isn't
    If FileName = vbNullString Then
    
        FileNameIsValid = False 'Invalid
        Exit Function
    
    Else
      
        'Loop through each "Bad Character" and test for an instance
        For i = 1 To Len(strBadChar)
            If InStr(FileName, Mid$(strBadChar, i, 1)) > 0 Then
                FileNameIsValid = False 'Invalid
                Exit For
            End If
        Next i
    
    End If
    
End Function
