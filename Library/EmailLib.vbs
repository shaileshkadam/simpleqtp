Option Explicit

''' #########################################################
''' <summary>
''' Send email using CDO/MAPI
''' </summary>
''' <remarks></remarks>	 
''' #########################################################

Class SendMailUsingCDO
 
    Private oMessage    'CDO.Message Object
    Private strFrom     'Sender's Email ID: XX@YY.COM
    Public Body         'Body Text from Text File

    ''' <summary>
    ''' Send Email Using CDO
    ''' </summary>
    ''' <param name="sEMailID" type="string">Sender's Mail ID String</param>
    ''' <param name="sPassword" type="string">Sender's Password String</param>
    ''' <param name="sTo" type="string">Recipient's Mail ID String (Primary)</param>
    ''' <param name="sCC" type="string">Recipient's Mail ID String (CC)</param>
    ''' <param name="sSubject" type="string">Subject String</param>
    ''' <param name="sBody" type="string">Body Message String</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Send(ByVal sEMailID, ByVal sPassword, ByVal sTo, ByVal sCC, ByVal sSubject, ByVal sBody)

        Dim oRegExp     'RegEx Object
        Dim sDetails    'Report Details
        Dim intStatus   'Report Status
        Dim sStepName   'Report Step
        
        'Sender ID has Class scope
        Me.From = sEmailID
        'Message Body
        If sBody <> "" Then Me.Body = sBody
 
        intStatus = micPass
        sStepName = " Sent"
 
        Set oRegExp = New RegExp
        oRegExp.Global = True
        oRegExp.Pattern = "<\w>|<\w\w>|<\w\d>"
        Set oMatches = oRegExp.Execute( Me.Body )
 
        'Build Message
        With oMessage
            .Subject = sSubject
            .From = sEmailID
            .To = sTo
            .CC = sCC
            'BCC Property can be added as well:
            '.BCC = sBCC
            'If HTML Tags found, use .HTMLBody
            If oMatches.Count > 0 Then
                .HTMLBody = Me.Body
            Else
                .TextBody = Me.Body
            End If
        End With
 
        Set oMatches = Nothing
        Set oRegExp = Nothing
 
        With oMessage.Configuration.Fields
            'Sender's Mail ID
            .Item("http://schemas.microsoft.com/cdo/configuration/" &_
            "sendusername") = sEmailID
            'Sender's Password
            .Item("http://schemas.microsoft.com/cdo/configuration/" &_
            "sendpassword") = sPassword
            'Name/IP of SMTP Server
            .Item("http://schemas.microsoft.com/cdo/configuration/" &_
            "smtpserver") = cdoSMTPServer
            'Server Port
            .Item("http://schemas.microsoft.com/cdo/configuration/" &_
            "smtpserverport") = cdoOutgoingMailSMTP
            'Send Using: (1) Local SMTP Pickup Service (2) Use SMTP Over Network
            .Item("http://schemas.microsoft.com/cdo/configuration/" &_
            "sendusing") = cdoSendUsing
            'Authentication Used: (1) None (2) Basic (3) NTLM
            .Item("http://schemas.microsoft.com/cdo/configuration/" &_
            "smtpauthenticate") = cdoAuthenticationType
            'SMTP Server Requires SSL/STARTTLS: Boolean
            .Item("http://schemas.microsoft.com/cdo/configuration/" &_
            "smtpusessl") = cdoUseSSL
            'Maximum Time in Seconds CDO will try to Establish Connection
            .Item("http://schemas.microsoft.com/cdo/configuration/" &_
            "smtpconnectiontimeout") = cdoTimeout
            'Update Configuration Entries
            .Update
        End With
 
        'Report Details
        sDetails = "SMTP Server: " & cdoSMTPServer & vbLf
        sDetails = sDetails & "Sender: " & sEMailID & vbLf
        sDetails = sDetails & "Recipient: " & sTo & vbLf
        sDetails = sDetails & "Server Port: " & cdoOutgoingMailSMTP & vbLf
        sDetails = sDetails & "SSL Used: " & cdoUseSSL & vbLf
        sDetails = sDetails & "Authentication Type: " & cdoAuthenticationType & vbLf
        sDetails = sDetails & "SMTP Service Type: " & cdoSendUsing & vbLf & vbLf
        sDetails = sDetails & "Subject: " & sSubject & vbLf & vbLf
        sDetails = sDetails & "Body: " & sBody
 
        On Error Resume Next
            'Send Message
            oMessage.Send
            If Err.Number <> 0 Then
                intStatus = micWarning
                sStepName = " Not Sent"
                sDetails = sDetails & vbLf & "Error Description: " & Err.Description
            End If
        On Error Goto 0
 
        'If you're not using QTP, please disable the statement below:
        'Reporter.ReportEvent intStatus, "EMail" & sStepName, sDetails
        
    End Function
 
    ''' <summary>
    ''' Loads Body message from a Text File
    ''' </summary>
	''' <param name="sCompleteFilePath" type="string">Complete Path to the Text File containing the Body Message</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LoadBodyMessage(ByVal sCompleteFilePath )

        CONST ForReading = 1 
        Dim oFSO, oFile
        Set oFSO = CreateObject( "Scripting.FileSystemObject" )
        Set oFile = oFSO.OpenTextFile( sCompleteFilePath, ForReading )
        Me.Body = oFile.ReadAll
        oFile.Close: Set oFile = Nothing
        Set oFSO = Nothing
        
    End Function
 
    ''' <summary>
    ''' Class_Initialize, Binds to the CDO Object
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Sub Class_Initialize
  
        Set oMessage = CreateObject( "CDO.Message" )
        
    End Sub
   
    ''' <summary>
    ''' Class_Terminate, Release the CDO Object
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Sub Class_Terminate

        Set oMessage = Nothing
        
    End Sub

    ''' <summary>
    ''' Readonly property configuration for SMTP Service
    ''' </summary>
    ''' <remarks></remarks>
    Private Property Get cdoSendUsing  'As Integer
   
        cdoSendUsing = 2    'Use SMTP Over The Network
        'cdoSendUsing = 1    'Use Local SMTP Service Pickup Directory
        
    End Property
 
    ''' <summary>
    ''' Maximum time in seconds CDO will try to establish a connection
    ''' </summary>
    ''' <remarks></remarks>
    Private Property Get cdoTimeout  'As Integer
 
        'cdoTimeout = 15    'Seconds
        cdoTimeout = 45    'Seconds
        'cdoTimeout = 75    'Seconds
        
    End Property
 
    ''' <summary>
    ''' Type of Authentication to be used
    ''' </summary>
    ''' <remarks></remarks>  
    Private Property Get cdoAuthenticationType  'As Integer
  
        'cdoAuthenticationType = 0    'No Authentication
        cdoAuthenticationType = 1    'Basic Authentication
        'cdoAuthenticationType = 2    'NTML Authentication
        
    End Property
   
    ''' <summary>
    ''' Server Port
    ''' </summary>
    ''' <remarks></remarks> 
    Private Property Get cdoOutgoingMailSMTP  'As Integer
  
        If InStr(1, Lcase(Me.From), "@gmail") <> 0 Then
            cdoOutgoingMailSMTP = 465
        ElseIf InStr(1, LCase(Me.From), "@aol") <> 0 Then
            cdoOutgoingMailSMTP = 587
        Else
            cdoOutgoingMailSMTP = 25
        End If
        
    End Property
 
    ''' <summary>
    ''' Name/IP of SMTP Server
    ''' </summary>
    ''' <remarks></remarks> 
    Private Property Get cdoSMTPServer  'As String

        If InStr(1, LCase(Me.From), "@yahoo") <> 0 Then
            cdoSMTPServer = "smtp.mail.yahoo.com"
        ElseIf InStr(1, LCase(Me.From), "@gmail") <> 0 Then
            cdoSMTPServer = "smtp.gmail.com"
        ElseIf InStr(1, LCase(Me.From), "@hotmail") <> 0 Or _
               InStr(1, LCase(Me.From), "@live") <> 0 Then
            cdoSMTPServer = "smtp.live.com"
        ElseIf InStr(1, LCase(Me.From), "@aol") <> 0 Then
            cdoSMTPServer = "smtp.aol.com"
        End If  
          
    End Property
       
    ''' <summary>
    ''' Setting for SMTP Server's use of SSL (Boolean)
    ''' </summary>
    ''' <remarks></remarks> 
    Private Property Get cdoUseSSL  'As Boolean

        cdoUseSSL = True
        If InStr(1, LCase(Me.From), "@aol") <> 0 Then
            cdoUseSSL = False
        End If
        
    End Property
 
    ''' <summary>
    ''' Sender's Email ID
    ''' </summary>
    ''' <remarks></remarks> 
    Public Property Let From( ByVal Val )
           
        strFrom = Val
           
    End Property
    
    Public Property Get From 'As String
    
        From = strFrom
        
    End Property
 
End Class
 
''' <summary>
''' Sends an Email Using CDO to a recipient
''' </summary>
''' <param name="EMailID" type="string">Sender's Mail ID String</param>
''' <param name="Password" type="string">Sender's Password String</param>
''' <param name="Recipient" type="string">Recipient's Mail ID String (Primary)</param>
''' <param name="CC" type="string">Recipient's Mail ID String (CC)</param>
''' <param name="Subject" type="string">Subject String</param>
''' <param name="Body" type="string">Body Message String</param>
''' <returns></returns>
''' <remarks></remarks>
''' <example>
''' SendEmailUsingCDO "lwfwind@gmail.com", "123458", "lwfwind@126.com", "", "Subject", "Hello, this is a test mail."
''' SendEmailUsingCDO "lwfwind@gmail.com", "123458", "lwfwind@126.com", "", "Subject", "<h1>Hello</h1><p>this is a html mail</p>"
''' </example>
Public Function SendEmailUsingCDO(ByVal EmailID, ByVal Password, ByVal Recipient, ByVal CC, ByVal Subject, ByVal Body)
	
	Dim oEmail
    Set oEmail = New SendMailUsingCDO
    With oEmail.Send EmailID, Password, Recipient, CC, Subject, Body
    End With
    
End Function

''' <summary>
''' Sends an Email Using CDO and Body Message will be from a specific text file
''' </summary>
''' <param name="EMailID" type="string">Sender's Mail ID String</param>
''' <param name="Password" type="string">Sender's Password String</param>
''' <param name="Recipient" type="string">Recipient's Mail ID String (Primary)</param>
''' <param name="CC" type="string">Recipient's Mail ID String (CC)</param>
''' <param name="Subject" type="string">Subject String</param>
''' <param name="sCompleteFilePath" type="string">Text File containing the Body Message</param>
''' <returns></returns>
''' <remarks></remarks>
''' <example>
''' EmailFromFileUsingCDO "lwfwind@gmail.com", "123458", "lwfwind@126.com", "", "Subject", "Hello, this is a test mail."
''' EmailFromFileUsingCDO "lwfwind@gmail.com", "123458", "lwfwind@126.com", "", "Subject", "<h1>Hello</h1><p>this is a html mail</p>"
''' </example>
Public Function EmailFromFileUsingCDO(ByVal EmailID,ByVal Password,ByVal Recipient,ByVal CC,ByVal Subject,ByVal sCompleteFilePath )

    Dim EmailFromFile
    Set EmailFromFile = New SendMailUsingCDO
    With EmailFromFile
        .LoadBodyMessage sCompleteFilePath
        .Send EmailID, Password, Recipient, CC, Subject, ""
    End with 
    
End Function


	
Class SendMailUsingMAPI

	Private ol
	
	''' <summary>
    ''' Class_Initialize, Binds to the CDO Object
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Sub Class_Initialize
  
        Set ol = WScript.CreateObject("Outlook.Application")
        
    End Sub
   
    ''' <summary>
    ''' Class_Terminate, Release the CDO Object
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Sub Class_Terminate

        Set ol = Nothing
        
    End Sub
	
	''' <summary>
    ''' Sends an Email Using using the Outlook client
    ''' </summary>
    ''' <param name="ToAddress" type="string">Recipient's Mail ID String (Primary)</param>
    ''' <param name="MessageSubject" type="string">Subject String</param>
    ''' <param name="MessageBody" type="string">Body Message</param>
    ''' <param name="MessageAttachment" type="string">Attachment</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function Send(ByVal ToAddress,ByVal MessageSubject, ByVal MessageBody, ByVal MessageAttachment)
	
		' connect to Outlook
		Dim ns, Message,arrToAddress,i
		Set ns = ol.getNamespace("MAPI")
		ns.logon "","",true,false
		Set Message = ol.CreateItem(olMailItem)
		Message.Attachments.Add(MessageAttachment)
		Message.Subject = MessageSubject
		Message.Body = MessageBody & vbCrLf
		' validate the recipient, just in case...
		arrToAddress = Split(ToAddress, ";")
		For i = LBound(arrToAddress) To UBound(arrToAddress)
			Set myRecipient = ns.CreateRecipient(Trim(arrToAddress(i)))
			myRecipient.Resolve
			If Not myRecipient.Resolved Then
				MsgBox "unknown recipient"
			Else
				Message.Recipients.Add(myRecipient)
			End If	
		Next
		Message.Send
		
	End Function

End Class        

''' <summary>
''' Sends an Email Using the Outlook client
''' </summary>
''' <param name="ToAddress" type="string">Recipient's Mail ID String (Primary)</param>
''' <param name="MessageSubject" type="string">Subject String</param>
''' <param name="MessageBody" type="string">Body Message</param>
''' <param name="MessageAttachment" type="string">Attachment</param>
''' <returns></returns>
''' <remarks></remarks>
''' <example>
''' SendEmailUsingMAPI "WLu@StateStreet.com;wflu@hengtiansoft.com", "VBS MAPI HowTo", "email via MAPI", "C:\Config.xls"
''' </example>
Public Function SendEmailUsingMAPI(ByVal ToAddress,ByVal MessageSubject, ByVal MessageBody, ByVal MessageAttachment)

	Dim oEmail
	Set oEmail = New SendMailUsingMAPI
	With oEmail
	    .Send ToAddress, MessageSubject, MessageBody, MessageAttachment
	End with 
	
End Function