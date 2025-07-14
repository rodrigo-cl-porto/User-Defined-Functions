Sub SendEmail(Sender As String, Recipient As String, Subject As String, Message As String, Optional CarbonCopy As String, Optional BlindCarbonCopy As String)

    Const CDO_DEFAULT_SETTINGS    As Integer = -1 'Default settings from the system or current profile.
    Const CDO_NTLM_AUTHENTICATION As Integer = 2 'Integrated Windows Authentication (NTLM). Used in corporate environments with Exchange Server.
    Const CDO_SEND_USING_PORT     As Integer = 2 'Send email directly via SMTP port
    Const CDO_SERVER_PORT         As Integer = 587 'Authenticated sending with STARTTLS
    Const CDO_SMTP_SERVER         As String = "mailhost.yourdomain.net"
    Dim Email         As Object: Set Email = CreateObject("CDO.Message")
    Dim EmailSettings As Object: Set EmailSettings = CreateObject("CDO.Configuration")

    EmailSettings.load CDO_DEFAULT_SETTINGS
    With EmailSettings.Fields
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = CDO_SEND_USING_PORT
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = CDO_SMTP_SERVER
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = CDO_NTLM_AUTHENTICATION
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = CDO_SERVER_PORT
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
