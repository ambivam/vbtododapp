Attribute VB_Name = "modNotifications"
Option Explicit

' Constants for SMTP settings
Private Const SMTP_SERVER As String = "smtp.gmail.com"
Private Const SMTP_PORT As Integer = 587
Private Const SMTP_USERNAME As String = "your_email@gmail.com"
Private Const SMTP_PASSWORD As String = "your_app_password"

' Constants for SMS API (using Twilio as example)
Private Const TWILIO_ACCOUNT_SID As String = "your_account_sid"
Private Const TWILIO_AUTH_TOKEN As String = "your_auth_token"
Private Const TWILIO_FROM_NUMBER As String = "your_twilio_number"

Public Function SendEmailAlert(ByVal ToEmail As String, ByVal Subject As String, ByVal Body As String) As Boolean
    On Error GoTo SendError
    
    Dim objEmail As Object
    Set objEmail = CreateObject("CDO.Message")
    
    With objEmail.Configuration.Fields
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTP_SERVER
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTP_PORT
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = SMTP_USERNAME
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SMTP_PASSWORD
        .Update
    End With
    
    With objEmail
        .From = SMTP_USERNAME
        .To = ToEmail
        .Subject = Subject
        .TextBody = Body
        .Send
    End With
    
    Set objEmail = Nothing
    SendEmailAlert = True
    Exit Function
    
SendError:
    SendEmailAlert = False
End Function

Public Function SendSMSAlert(ByVal ToNumber As String, ByVal Message As String) As Boolean
    On Error GoTo SendError
    
    ' Using Windows HTTP object to make API call to Twilio
    Dim objHTTP As Object
    Set objHTTP = CreateObject("MSXML2.XMLHTTP")
    
    ' Twilio API endpoint
    Dim url As String
    url = "https://api.twilio.com/2010-04-01/Accounts/" & TWILIO_ACCOUNT_SID & "/Messages.json"
    
    ' Prepare form data
    Dim postData As String
    postData = "To=" & ToNumber & "&From=" & TWILIO_FROM_NUMBER & "&Body=" & Message
    
    ' Send request
    objHTTP.Open "POST", url, False
    objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objHTTP.setRequestHeader "Authorization", "Basic " & Base64Encode(TWILIO_ACCOUNT_SID & ":" & TWILIO_AUTH_TOKEN)
    objHTTP.send postData
    
    SendSMSAlert = (objHTTP.Status = 201)
    Set objHTTP = Nothing
    Exit Function
    
SendError:
    SendSMSAlert = False
End Function

Private Function Base64Encode(ByVal Text As String) As String
    Dim objXML As Object
    Dim objNode As Object
    
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    
    objNode.DataType = "bin.base64"
    objNode.Text = StrConv(Text, vbFromUnicode)
    Base64Encode = objNode.Text
    
    Set objNode = Nothing
    Set objXML = Nothing
End Function

Public Sub SendTaskReminder(ByVal TaskText As String, ByVal DueDate As String, _
                          ByVal UserEmail As String, ByVal UserPhone As String)
    Dim emailSubject As String
    Dim emailBody As String
    Dim smsMessage As String
    
    emailSubject = "Todo Reminder: Task Due Soon"
    emailBody = "Your task '" & TaskText & "' is due on " & DueDate & "." & vbCrLf & vbCrLf & _
                "Please complete it as soon as possible."
                
    smsMessage = "Todo Reminder: '" & TaskText & "' is due on " & DueDate
    
    ' Send email notification
    SendEmailAlert UserEmail, emailSubject, emailBody
    
    ' Send SMS notification
    If Trim(UserPhone) <> "" Then
        SendSMSAlert UserPhone, smsMessage
    End If
End Sub
