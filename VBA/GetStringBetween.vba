Function GetStringBetween(str As String, startStr As String, endStr As String) As String

    Dim regex As Object
    Dim matches As Object
    Dim pattern As String

    pattern = "(?<=" & startStr & "\s*).*?(?=\s*" & endStr & ")"

    Set regex = CreateObject("VBScript.RegExp")
    With regex
      .Pattern = pattern
      .IgnoreCase = True
      .Global = False

      If .Test(str) Then
          Set matches = .Execute(str)
          GetStringBetween = matches(0).Value
      Else
          GetStringBetween = ""
      End If
    End With

End Function
