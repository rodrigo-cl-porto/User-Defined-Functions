Sub SendEmail(Sender As String, Recipient As String, Subject As String, Message As String, Optional CarbonCopy As String, Optional BlindCarbonCopy As String)

    Const LOAD_DEFAULT_CONFIGURATION As Long = -1
    Const cdoNtlmAuthentication as Integer = 2 'Integrated Windows Authentication (NTLM). Used in corporate environments with Exchange Server.
    Const cdoSendUsingPort      as Integer = 2 'Send email directly via SMTP port
    Const cdoServerPort         as Integer = 587 'Authenticated sending with STARTTLS
    Dim Email         As Object: Set Email = CreateObject("CDO.Message")
    Dim EmailSettings As Object: Set EmailSettings = CreateObject("CDO.Configuration")

    EmailSettings.Load LOAD_DEFAULT_CONFIGURATION
    With EmailSettings.Fields
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mailhost.yourdomain.net"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoNtlmAuthentication
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = cdoServerPort
        '.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        .Update
    End With
    
    With Email
        Set .Configuration = EmailSettings
        .From = Sender
        .To = Recipient
        .CC = CarbonCopy
        .BCC = BlindCarbonCopy
        .Subject = Subject
        .HtmlBody = Message
        .Send
    End With
    
End Sub
