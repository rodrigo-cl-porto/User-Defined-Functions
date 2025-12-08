# [`SendEmail`](/src/vba/SendEmail.vba)

Sends an HTML email using CDO (Collaboration Data Objects) with NTLM authentication, typically used in corporate environments with Exchange Server.

## Syntax

```vb
SendEmail( _
    Sender As String, _
    Recipient As String, _
    Subject As String, _
    Message As String, _
    Optional CarbonCopy As String, _
    Optional BlindCarbonCopy As String _
)
```

## Parameters

- `Sender`: Email address of the sender
- `Recipient`: Email address(es) of the recipient(s)
- `Subject`: Subject line of the email
- `Message`: HTML-formatted body of the email
- `CarbonCopy`: (_optional_) Email address(es) for CC recipients
- `BlindCarbonCopy`: (_optional_) Email address(es) for BCC recipients

## Remarks

- Uses CDO with NTLM authentication (Windows Authentication)
- Configured for SMTP with STARTTLS (port 587)
- Supports HTML formatting in the message body
- Multiple recipients can be specified using semicolon (;) as separator
- No explicit error handling is implemented

## **Configuration Constants**

- `CDO_DEFAULT_SETTINGS`: -1 (Use system default settings)
- `CDO_NTLM_AUTHENTICATION`: 2 (Windows Authentication)
- `CDO_SEND_USING_PORT`: 2 (Direct SMTP)
- `CDO_SERVER_PORT`: 587 (STARTTLS port)
- `CDO_SMTP_SERVER`: "mailhost.yourdomain.net" (SMTP server address)

## Dependencies

- Requires CDO to be available on the system
- Requires proper SMTP server configuration
- Requires appropriate network/firewall access

## Example

```vb
Call SendEmail( _
    "sender@company.com", _
    "recipient@company.com", _
    "Test Subject", _
    "<h1>Hello</h1><p>This is a test email.</p>", _
    "cc@company.com", _
    "bcc@company.com")
```
