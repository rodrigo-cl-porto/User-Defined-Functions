Sub SendEmail(Sender As String, Recipient As String, Subject As String, Message As String, Optional CarbonCopy As String, Optional BlindCarbonCopy As String)
    
    Dim Email         As Object: Set Email = CreateObject("CDO.Message")
    Dim EmailSettings As Object: Set EmailSettings = CreateObject("CDO.Configuration")

    CDO_Config.Load -1
    With EmailSettings.Fields
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mailhost.subsea7.net"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        '.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "xxxx"
        '.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "xxxx"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
        .Update
    End With
    
    With Email
        Set .Configuration = EmailSettings
        .From = Sender
        .To = Recipient
        .CC = CarbonCopy
        .BCC = BlindCarbonCopy
        .Subject = Subject
        .htmlBody = Message
        .Send
    End With
    
End Sub
