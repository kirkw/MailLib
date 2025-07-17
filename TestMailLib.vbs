' MailLib Test Script
' This script demonstrates how to use the MailLib.EmailSender class
' 
' PRIVATE VERSION - Keep this file private and never commit credentials
' PUBLIC VERSION - Uses environment variable support

Option Explicit

' Configuration - PRIVATE VERSION (replace with your actual values)
Dim SMTP_HOST, SMTP_PORT, SMTP_USERNAME, SMTP_PASSWORD, FROM_EMAIL, TO_EMAIL

' ===========================================
' PRIVATE CONFIGURATION - REPLACE WITH YOUR VALUES
' ===========================================
SMTP_HOST = ""           ' e.g., "smtp.gmail.com"
SMTP_PORT = 587          ' e.g., 587 for STARTTLS, 465 for SSL
SMTP_USERNAME = ""       ' e.g., "your-email@gmail.com"
SMTP_PASSWORD = ""       ' e.g., "your-app-password"
FROM_EMAIL = ""          ' e.g., "sender@yourdomain.com"
TO_EMAIL = ""            ' e.g., "recipient@example.com"

' ===========================================
' PUBLIC VERSION USING ENVIRONMENT VARIABLES
' ===========================================
' Uncomment the lines below to use environment variables instead
' This makes the script safe to share publicly
'
 SMTP_HOST     = GetEnvironmentVariable("MAILLIB_SMTP_HOST")
 SMTP_PORT     = GetEnvironmentVariable("MAILLIB_SMTP_PORT")
 SMTP_USERNAME = GetEnvironmentVariable("MAILLIB_SMTP_USERNAME")
 SMTP_PASSWORD = GetEnvironmentVariable("MAILLIB_SMTP_PASSWORD")
 FROM_EMAIL    = GetEnvironmentVariable("MAILLIB_FROM_EMAIL")
 TO_EMAIL      = GetEnvironmentVariable("MAILLIB_TO_EMAIL")

' ===========================================
' TEST FUNCTION
' ===========================================
Sub TestMailLib()
    On Error Resume Next
    
    ' Validate configuration
    If SMTP_HOST = "" Then
        WScript.Echo "ERROR: SMTP_HOST is not set. Please configure the script."
        WScript.Quit 1
    End If
    
    If SMTP_USERNAME = "" Then
        WScript.Echo "ERROR: SMTP_USERNAME is not set. Please configure the script."
        WScript.Quit 1
    End If
    
    If FROM_EMAIL = "" Then
        WScript.Echo "ERROR: FROM_EMAIL is not set. Please configure the script."
        WScript.Quit 1
    End If
    
    If TO_EMAIL = "" Then
        WScript.Echo "ERROR: TO_EMAIL is not set. Please configure the script."
        WScript.Quit 1
    End If
    
    WScript.Echo "Starting MailLib Test..."
    WScript.Echo "SMTP Server: " & SMTP_HOST & ":" & SMTP_PORT
    WScript.Echo "From: " & FROM_EMAIL
    WScript.Echo "To: " & TO_EMAIL
    WScript.Echo ""
    
    ' Create EmailSender object
    Dim EmailSender
    Set EmailSender = CreateObject("MailLib.EmailSender")
    
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Failed to create MailLib.EmailSender object."
        WScript.Echo "Error: " & Err.Number & " = " & Err.Description
        WScript.Echo "Make sure MailLib.dll is registered with:"
        WScript.Echo "RegAsm.exe MailLib.dll /codebase /tlb"
        WScript.Quit 1
    End If
    
    ' Configure SMTP settings
    EmailSender.Host = SMTP_HOST
    EmailSender.Port = SMTP_PORT
    EmailSender.UserName = SMTP_USERNAME
    EmailSender.Password = SMTP_PASSWORD
    
    ' Set connection type based on port
    If SMTP_PORT = 465 Then
        ' SSL on connect
        EmailSender.ConnectionType = 2  ' ServerSecurity.Ssl
    ElseIf SMTP_PORT = 587 Then
        ' STARTTLS
        EmailSender.ConnectionType = 3  ' ServerSecurity.Tls
    Else
        ' Auto-detect
        EmailSender.ConnectionType = 1  ' ServerSecurity.Auto
    End If
    
    ' Set message details
    EmailSender.From = FROM_EMAIL
    EmailSender.To = TO_EMAIL
    EmailSender.Subject = "MailLib Test Email - " & Now()
    
    ' Add message body
    EmailSender.AppendToHtmlBody "<h2>MailLib Test Email</h2>"
    EmailSender.AppendToHtmlBody "<p>This is a test email sent using <strong>MailLib</strong>.</p>"
    EmailSender.AppendToHtmlBody "<p>Sent at: " & Now() & "</p>"
    EmailSender.AppendToHtmlBody "<p>SMTP Server: " & SMTP_HOST & ":" & SMTP_PORT & "</p>"
    EmailSender.AppendToHtmlBody "<hr>"
    EmailSender.AppendToHtmlBody "<p><em>This email was sent using the MailLib wrapper for MailKit.</em></p>"
    
    ' Add text version as well
    EmailSender.AppendToTextBody "MailLib Test Email" & vbCrLf & vbCrLf
    EmailSender.AppendToTextBody "This is a test email sent using MailLib." & vbCrLf
    EmailSender.AppendToTextBody "Sent at: " & Now() & vbCrLf
    EmailSender.AppendToTextBody "SMTP Server: " & SMTP_HOST & ":" & SMTP_PORT & vbCrLf
    
    ' Send the email
    WScript.Echo "Testing Call..."
	Call EmailSender.ClearBccAddresses
   If Err.Number <> 0 Then
        WScript.Echo "ERROR: Failed to clear BccAddresses."
        WScript.Echo "Error: " & Err.Description
        WScript.Echo "Error Number: " & Err.Number 
		WScript.Echo "Error Source: " & Err.Source
        WScript.Quit 1
    End If

    WScript.Echo "Sending email..."
    Dim success
    success = EmailSender.Send() ' returns bool
    
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Failed to send email."
        WScript.Echo "Error: " & Err.Description
        WScript.Echo "Error Number: " & Err.Number 
		WScript.Echo "Error Source: " & Err.Source
        WScript.Quit 1
    End If
    
    If success Then
        WScript.Echo "SUCCESS: Email sent successfully!"
        WScript.Echo "Check the recipient's inbox for the test email."
    Else
        WScript.Echo "ERROR: Email sending failed."
        WScript.Echo "Check the Windows Event Log for more details."
    End If
    
    ' Clean up
    Set EmailSender = Nothing
    
    WScript.Echo ""
    WScript.Echo "Test completed."
End Sub

' ===========================================
' ENVIRONMENT VARIABLE HELPER FUNCTION
' ===========================================
Function GetEnvironmentVariable(variableName)
    Dim objShell
    Set objShell = CreateObject("WScript.Shell")
    GetEnvironmentVariable = objShell.Environment("PROCESS")(variableName)
    Set objShell = Nothing
End Function

' ===========================================
' CDO CONFIGURATION TEST (Alternative approach)
' ===========================================
Sub TestMailLibWithCDOConfig()
    On Error Resume Next
    
    WScript.Echo "Testing MailLib with CDO-style configuration..."
    
    ' Create CDO-style configuration
    Dim config
    Set config = CreateObject("CDO.Configuration")
    
    ' Set CDO configuration fields
    config.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    config.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTP_HOST
    config.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTP_PORT
    config.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False ' was true
    config.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = SMTP_USERNAME
    config.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SMTP_PASSWORD
    config.Fields("http://schemas.microsoft.com/cdo/configuration/replyto") = FROM_EMAIL  ' Shockingly this works!
    config.Fields("http://schemas.microsoft.com/cdo/configuration/from") = FROM_EMAIL
    
    ' Update the configuration
    config.Fields.Update
	
	'Wscript.Echo "Starting Loop ==============="
	'Dim i
	'For i = 0 To config.Fields.Count - 1
	'	On Error Resume Next
	'	WScript.Echo "Field " & i & ": " & config.Fields(i).Name & " = " & config.Fields(i).Value
	'	On Error Goto 0
	'Next
	
    ' Set the configuration on EmailSender
    ' Create EmailSender object
    Dim EmailSender
    Set EmailSender = CreateObject("MailLib.EmailSender")
    
    EmailSender.Configuration = config
	EmailSender.Log_File = "D:\_email_sender.log"
    
    ' Set message details
    Wscript.Echo "From   : " & EmailSender.From ' = FROM_EMAIL
    Wscript.Echo "ReplyTo: " & EmailSender.ReplyTo  '= TO_EMAIL
	EmailSender.ReplyTo = EmailSender.ReplyTo
	EmailSender.To  = TO_EMAIL
    EmailSender.Subject = "MailLib CDO Config Test - " & Now()
    EmailSender.AppendToHtmlBody "<h2>MailLib CDO Configuration Test</h2>"
    EmailSender.AppendToHtmlBody "<p>This email was sent using CDO-style configuration." & Now() & "</p>"
    
    ' Send the email
    Dim success
    success = EmailSender.Send()
    
    If success Then
        WScript.Echo "SUCCESS: CDO configuration test email sent!"
    Else
        WScript.Echo "ERROR: CDO configuration test failed."
    End If
    
    Set EmailSender = Nothing
    Set config = Nothing
End Sub

' ===========================================
' MAIN EXECUTION
' ===========================================
WScript.Echo "MailLib Test Script"
WScript.Echo "=================="
WScript.Echo ""

' Run the main test
'TestMailLib()

WScript.Echo ""
WScript.Echo "Press any key to run CDO configuration test..."
WScript.StdIn.Read(1)

' Run the CDO configuration test
TestMailLibWithCDOConfig()

WScript.Echo ""
WScript.Echo "All tests completed. Press any key to exit..."
WScript.StdIn.Read(1) 