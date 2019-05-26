<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: clsMail.asp
'  Date:     $Date: 2004/03/11 06:29:08 $
'  Version:  $Revision: 1.2 $
'  Purpose:  Class object to assist when emailing notifications
' ----------------------------------------------------------------------------------

Class clsMail
  ' Class Mail - Provides a common interface for sending e-mail
  ' Supported Methods
  ' 1 - CDOSYS(smtp)
  ' 2 - CDONTS
  '     * Local IIS SMTP/Exchange only
  '     * Does not support authentication
  ' 3 - JMail
  ' 4 - ASPMail
  '     * Does not support HTML
  '     * Does not support authentication
  ' 5 - ASPEmail
  '     * Premium version requried for authentication

  ' Example (using CDOSYS):
  ' Set oMail = New clsMail
  ' oMail.MailType = 1
  ' oMail.RemoteHost = "smtp.company.com"
  ' oMail.From = "user@company.com"
  ' oMail.Subject = "Message Subject"
  ' oMail.Body = "This is the message body."
  ' oMail.AddRecipient "user2@other.com"
  ' oMail.AddRecipient "user3@other.com"
  ' If Not oMail.Send Then
  '   Response.Write oMail.LastError
  ' End If
  '
  ' -- To send the email in HTML --
  ' oMail.UseHTML = True
  '
  ' -- To use SMTP authentication --
  ' oMail.User = "username"
  ' oMail.Password = "password"


  ' ##################
  ' Private Properties
  ' ##################
  Private m_Mail, m_MailMethod, m_LastError
  Private m_From, m_HTML, m_Subject, m_Body
  Private m_SMTPServer, m_SMTPUser, m_SMTPPassword
  Private m_ToAddr


  ' ##################
  ' Public Properties
  ' ##################
  Public Property Let MailMethod(ByRef f_MailMethod)
    m_MailMethod = f_MailMethod
  End Property

  Public Property Let From(ByRef f_From)
    m_From = Trim(f_From)
  End Property

  Public Property Let UseHTML(ByRef f_HTML)
    If f_HTML Then
      m_HTML = True
    Else
      m_HTML = False
    End If
  End Property

  Public Property Let Subject(ByRef f_Subject)
    m_Subject = Trim(f_Subject)
  End Property

  Public Property Let Body(ByRef f_Body)
    m_Body = Trim(f_Body)
  End Property

  Public Property Let RemoteHost(ByRef f_Server)
    m_SMTPServer = Trim(f_Server)
  End Property

  Public Property Let User(ByRef f_User)
    m_SMTPUser = Trim(f_User)
  End Property

  Public Property Let Password(ByRef f_Password)
    m_SMTPPassword = f_Password
  End Property

  Public Property Get LastError() ' As String
    LastError = m_LastError
  End Property

  ' ##################
  ' Public Methods
  ' ##################
  Public Sub AddRecipient(ByVal f_Addr)
    f_Addr = Trim(f_Addr)
    If Len(f_Addr) > 0 Then
      Dim intCount
      intCount = m_ToAddr.Count + 1
      m_ToAddr.Add intCount, f_Addr
    End If
  End Sub

Public Function Send()  ' As Boolean

    Dim oMail, blnError
	Dim iConf 
	Dim Flds 
	Dim c
	Dim c1
	Dim c2
	Dim strJmailSend

	blnError = False
    
    On Error Resume Next
    
    Select Case m_MailMethod
    
		Case 1  ' CDOSYS(smtp)

			Set oMail = Server.CreateObject("CDO.Message")
			Set iConf = Server.CreateObject("CDO.Configuration")
			Set Flds = iConf.Fields

			With Flds
				.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'cdoSendUsingPort
				.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = m_SMTPServer 
				.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
				.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 20
				.Update
			End With

			If Len(m_SMTPUser) > 0 Then

				With Flds
				.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 ' cdoBasic
				.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = m_SMTPUser
				.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = m_SMTPPassword
				.Update
				End With

			End If

			Set oMail.Configuration = iConf
			oMail.From = m_From
			oMail.To = GetToAddrs
			oMail.Subject = m_Subject

			If m_HTML Then
				oMail.HTMLBody = m_Body
			Else
				oMail.TextBody = m_Body
			End If
			
		Case 2  ' CDONTS(local)
			Set oMail = Server.CreateObject("CDONTS.NewMail")
	
			If m_HTML Then
				oMail.BodyFormat = 0 ' cdoBodyFormatHTML
				oMail.MailFormat = 0 ' MIME
			Else
				oMail.BodyFormat = 1 ' cdoBodyFormatText
			End If
			oMail.To = GetToAddrs
			oMail.From = m_From
			oMail.Subject = m_Subject
			oMail.Body = m_Body

		Case 3  ' JMail
			Set oMail = Server.CreateObject("Jmail.Message")

			oMail.Logging = True
			oMail.ISOEncodeHeaders = False
			oMail.From = m_From
			oMail.Subject = m_Subject
		
			For c = 1 to m_ToAddr.Count
				oMail.AddRecipient m_ToAddr.Item(c)
			Next
		
			If m_HTML Then
				oMail.HTMLBody = m_Body
			Else
				oMail.Body = m_Body
			End If
		
			If Len(m_SMTPUser) > 0 Then
				oMail.MailServerUserName = m_SMTPUser
				oMail.MailServerPassWord = m_SMTPPassword
			End If
			
		Case 4  ' ASPMail
			Set oMail = Server.CreateObject("SMTPsvg.Mailer")

			oMail.RemoteHost = m_SMTPServer
			oMail.FromAddress = m_From
			oMail.Subject = m_Subject
			oMail.BodyText = m_Body

			For c1 = 1 to m_ToAddr.Count
				oMail.AddRecipient m_ToAddr.Item(c1), m_ToAddr.Item(c1)
			Next

		Case 5  ' ASPEmail
			Set oMail = Server.CreateObject("Persits.MailSender")

			oMail.Host = m_SMTPServer
			oMail.From = m_From
			oMail.Subject = m_Subject
			oMail.Body = m_Body
			oMail.IsHTML = m_HTML

			For c2 = 1 to m_ToAddr.Count
				oMail.AddAddress m_ToAddr.Item(c2)
			Next

			'If Len(m_STMPUser) > 0 Then
			'  oMail.Username = m_SMTPUser
			'  oMail.Password = m_SMTPPassword
			'End If

		Case Else
			blnError = True
			m_LastError = "Invalid mail type selected."
			
		End Select
		
		If Err.Number <> 0 Then

			Set oMail = Nothing
			blnError = True
			m_LastError = Err.Source & ": " & Err.Description

		Else

			Select Case m_MailMethod
			
				Case 1  ' CDOSYS(smtp)
					oMail.Send
					
				Case 2  ' CDONTS(local)
					oMail.Send
					
				Case 3  ' JMail
					strJmailSend = m_SMTPServer
					If Len(m_SMTPUser) > 0 Then
						strJmailSend = m_SMTPUser & ":" & m_SMTPPassword & "@" & strJmailSend
					End If            
					If Not oMail.Send(strJmailSend) Then
						blnError = True
						m_LastError = "JMail Error: " & oMail.Log
					End If
					
				Case 4  ' ASPMail
					If Not oMail.SendMail Then
						blnError = True
						m_LastError = "ASPMail Error: " & oMail.Response
					End If
					
				Case 5  ' ASPEmail
					oMail.Send
					
				Case Else
				
			End Select
			
			Set oMail = Nothing
			
			If Err.Number <> 0 Then
				blnError = True
				m_LastError = Err.Source & ": " & Err.Description
			End If
			
		End If
		
	Send = Not blnError
	
End Function


  ' ##################
  ' Private Methods
  ' ##################
  
Private Function ParseEMail( ByVal strText, ByVal lngCaseID )

	Dim objCase
	
	
	Set objCase = New clsCase
	
	objCase.ID = lngCaseID
	
	If Not objCase.Load Then
	
		objCase.LastError
	
	Else
	
		strText = Replace(strText, "[problemid]", objCase.ID)
		strText = Replace(strText, "[title]", objCase.Title)
		strText = Replace(strText, "[description]", objCase.Description)
		strText = Replace(strText, "[status]", objCase.Status.ItemName)
		strText = Replace(strText, "[RaisedDate]", DisplayDate(objCase.RaisedDate))
		strText = Replace(strText, "[ClosedDate]", DisplayDate(objCase.ClosedDate))
		strText = Replace(strText, "[category]", objCase.Cat.CatName)
		strText = Replace(strText, "[department]", objCase.Dept.DeptName)
		strText = Replace(strText, "[phone]", objCase.Contact.OfficePhone)
		strText = Replace(strText, "[location]", objCase.Contact.Location)
		strText = Replace(strText, "[solution]", objCase.Solution)
		strText = Replace(strText, "[baseurl]", Application("BASE_URL"))
		strText = Replace(strText, "[contactusername]", objCase.Contact.UserName)
		strText = Replace(strText, "[contactfullname]", objCase.Contact.FName & " " & objCase.Contact.LName)
		strText = Replace(strText, "[contactemail]", objCase.Contact.Email)
		strText = Replace(strText, "[repusername]", objCase.Rep.UserName)
		strText = Replace(strText, "[repfullname]", objCase.Rep.FName & " " & objCase.Rep.LName)
		strText = Replace(strText, "[repemail]", objCase.Rep.Email)
		strText = Replace(strText, "[url]", Application("BASE_URL") & "/casemodify.asp?id=" & objCase.ID)

	End If

	Set objCase = Nothing
	
	ParseEMail = strText	' Return the parsed string

End Function
  
  Private Function GetToAddrs() ' As String
    Dim c, strAddrs, arrAddr
    arrAddr = m_ToAddr.Items
    For c = 0 to m_ToAddr.Count -1
      If c > 0 Then
        strAddrs = strAddrs & ", "
      End If
      strAddrs = strAddrs & arrAddr(c) & " <" & arrAddr(c) & ">"
    Next
    GetToAddrs = strAddrs   
  End Function

  Private Sub Class_Initialize()
    m_HTML = False
    Set m_ToAddr = Server.CreateObject("Scripting.Dictionary")
  End Sub

  Private Sub Class_Terminate()
    Set m_ToAddr = Nothing
  End Sub
End Class
%>