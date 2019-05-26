<!--METADATA TYPE="TypeLib" NAME="Microsoft ActiveX Data Objects 2.5 Library" UUID="{00000205-0000-0010-8000-00AA006D2EA4}" VERSION="2.5"-->
<SCRIPT LANGUAGE="VBScript" RUNAT="SERVER">

' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: Public.asp
'  Date:     $Date: 2004/03/30 06:11:39 $
'  Version:  $Revision: 1.13 $
'  Purpose:  Contains an array a functions and procedures to assist in the running
'			 of the Liberum system
' ----------------------------------------------------------------------------------

' ################################
' PAGE START CODE
' ################################
' Declare Constants
'CONST adOpenStatic = 3
'CONST adOpenForwardOnly = 0

'CONST lhdDateOnly = 0
'CONST lhdDateTime = 1

'CONST lhdAddSQLDelim = 0
'CONST lhdNoSQLDelim = 1


' At the moment we are only allow for 32-Bit Security

CONST VIEW_USER = &H0000000000000001	
CONST VIEW_TECH = &H0000000000000002	
CONST VIEW_ADMIN = &H0000000000000004	
'CONST SPARE = &H0000000000000008

'CONST SPARE = &H0000000000000010
CONST PERM_CREATE_OWN = &H0000000000000020	
CONST PERM_READ_OWN = &H0000000000000040	
CONST PERM_MODIFY_OWN = &H0000000000000080	

CONST PERM_READ_GROUP = &H0000000000000100	
CONST PERM_MODIFY_GROUP = &H0000000000000200	
CONST PERM_READ_ASSIGNED =  &H0000000000000400
CONST PERM_MODIFY_ASSIGNED = &H0000000000000800

CONST PERM_CREATE_ALL = &H0000000000001000	
CONST PERM_READ_ALL = &H0000000000002000	
CONST PERM_MODIFY_ALL = &H0000000000004000	
'CONST SPARE = &H0000000000008000

CONST PERM_ACCESS_ADMIN = &H0000000000010000	
CONST PERM_ACCESS_TECH = &H0000000000020000	
CONST PERM_ACCESS_USER = &H0000000000040000	
'CONST SPARE = &H0000000000080000

CONST PERM_REOPEN_CASES = &H0000000000100000
CONST PERM_ACCESS_REPORTS = &H0000000000200000
'CONST SPARE = &H0000000000400000
'CONST SPARE = &H0000000000800000

CONST PERM_KB_READ = &H0000000001000000
CONST PERM_KB_MODIFY = &H0000000002000000
CONST PERM_KB_CREATE = &H0000000004000000
'CONST SPARE = &H0000000008000000

'CONST SPARE = &H0000000010000000
'CONST SPARE = &H0000000020000000
'CONST SPARE = &H0000000040000000
'CONST SPARE = &H0000000080000000


' For boolean values: SQL - lhd_True = 1 & lhd_False = 0
'					          : Access - lhd_True = -1 & lhd_False = 0
CONST lhd_True = -1
CONST lhd_False = 0


'CONST STATUS_OPEN
'CONST STATUS_CANCELLED
'CONST STATUS_CLOSED


' Set the public constants for the authentication type

Const lhd_ADAuthentication = 1
Const lhd_DBAuthentication = 2


'#################################


' Returns a ADO Connection object

Public Function CreateConnection()

	Dim strConnection, cnnDB
	
	' Check for usage of SQL securing or integrated security and use the correct
	' connection string.  Connection strings with DRIVER are ODBC, those with PROVIDER
	' are OLE DB connections.
	
	Select Case Application("DBType")
		Case 1	' SQL Sec
			strConnection = "Provider=SQLOLEDB.1;Data Source=" & Application("SQLServer")
			strConnection = strConnection & ";Initial Catalog=" & Application("SQLDBase")
			strConnection = strConnection & ";uid=" & Application("SQLUser") & ";pwd=" & Application("SQLPass")

		Case 2	' SQL Integrated Sec
			strConnection = "Provider=SQLOLEDB.1;Data Source=" & Application("SQLServer")
			strConnection = strConnection & ";Initial Catalog=" & Application("SQLDBase") & ";Integrated Security=SSPI"

		Case 3	' Access
			strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application("AccessPath")
			
		Case 4	' DSN
			strConnection = "DSN=" & Application("DSN_Name")

	End Select

	' Keep Errors from occuring

	If Not Application("Debug") Then
		On Error Resume Next
	End If

	' Create and open the database connection and save it as a session variable
	
	Set cnnDB = Server.CreateObject("ADODB.Connection")
	cnnDB.Open strConnection

	Set CreateConnection = cnnDB

	' Trap any errors

	If Err.Number <> 0 Then
'		Call TrapError(Err.Number, Err.Description, Err.Source)
	End If
	
End Function


' --------------------------------------------------------------------------
' 
' Routine: Sub CacheLanguageStrings
'
' Purpose: Caches all language labels and associated texts for speedy use
'		   on the asp pages.  The language to be cached is passed into the
'		   function via the "intLangID" variable.
'
' --------------------------------------------------------------------------
Public Sub CacheLanguageStrings(ByVal intLangID)

	Dim rstLabels
	Dim strSQL
	
	
	Set rstLabels = Server.CreateObject("ADODB.Recordset")
	
  strSQL = "SELECT tblLanguageLabels.LangLabel, tblLanguageTexts.LangText FROM tblLanguageTexts " &_
			 "INNER JOIN tblLanguageLabels ON tblLanguageTexts.LangLabelFK=tblLanguageLabels.LangLabelPK " &_
			 "WHERE tblLanguageTexts.LangFK=" & intLangID
			 
	rstLabels.Open strSQL, cnnDB
						    
    If rstLabels.BOF And rstLabels.EOF Then
						    
		' Raise Error, Labels didn't load
							
	Else
	
		While Not rstLabels.EOF
	
			Application("lhd_" & rstLabels.Fields("LangLabel")) = rstLabels.Fields("LangText")
				
			rstLabels.MoveNext
		
		WEnd

	End If

	rstLabels.Close 
	Set rstLabels = Nothing

End Sub


' --------------------------------------------------------------------------
' 
' Routine: Sub ClearLanguageStrings
'
' Purpose: Cleans out all the Cached language labels and associated texts.
'
' --------------------------------------------------------------------------
Public Sub ClearLanguageStrings()

	Dim rstLabels
	
	
	Set rstLabels = Server.CreateObject("ADODB.Recordset")
	
    strSQL = "SELECT LangLabel FROM tblLanguageLabels"
			 
	rstLabels.Open strSQL, cnnDB
						    
    If rstLabels.BOF And rstLabels.EOF Then
						    
		' Raise Error, Labels didn't load
							
	Else
	
		While Not rstLabels.EOF
	
			Application("lhd_" & rstLabels.Fields("LangLabel")) = Empty
				
			rstLabels.MoveNext
		
		WEnd

	End If

	rstLabels.Close 
	Set rstLabels = Nothing

End Sub


Public Sub SetApplicationParams()

	' Load Systems Parameters into Application variables
	
	Dim objParam


	Set objParam = New clsParameter
	
	Application("AUTH_TYPE") = CInt(objParam.GetValue("AUTH_TYPE"))
	Application("DATE_FORMAT") = objParam.GetValue("DATE_FORMAT")
	Application("DEFAULT_STATUS") = CInt(objParam.GetValue("DEFAULT_STATUS"))
	Application("DEFAULT_PRIORITY") = CInt(objParam.GetValue("DEFAULT_PRIORITY"))
	Application("DEFAULT_LANGUAGE") = CInt(objParam.GetValue("DEFAULT_LANGUAGE"))
	Application("DEFAULT_ROLE") = CInt(objParam.GetValue("DEFAULT_ROLE"))
	Application("ENABLE_ATTACHMENTS") = CInt(objParam.GetValue("ENABLE_ATTACHMENTS"))
	Application("ENABLE_EMAIL") = CInt(objParam.GetValue("ENABLE_EMAIL"))
	Application("ENABLE_INOUT") = CInt(objParam.GetValue("ENABLE_INOUT"))
	Application("ENABLE_KB") = CInt(objParam.GetValue("ENABLE_KB"))
	Application("ENABLE_REPORTS") = CInt(objParam.GetValue("ENABLE_REPORTS"))
	Application("EMAIL_METHOD") = CInt(objParam.GetValue("EMAIL_METHOD"))
	Application("MAX_ATTACHMENT_SIZE") = CLng(objParam.GetValue("MAX_ATTACHMENT_SIZE"))
	Application("ITEMS_PER_PAGE") = CInt(objParam.GetValue("ITEMS_PER_PAGE"))
	Application("SITE_NAME") = objParam.GetValue("SITE_NAME")
	Application("STATUS_OPEN") = CInt(objParam.GetValue("STATUS_OPEN"))
	Application("STATUS_CANCELLED") = CInt(objParam.GetValue("STATUS_CANCELLED"))
	Application("STATUS_CLOSED") = CInt(objParam.GetValue("STATUS_CLOSED"))
	Application("SYSTEM_EMAIL") = objParam.GetValue("SYSTEM_EMAIL")
	Application("TIME_FORMAT") = objParam.GetValue("TIME_FORMAT")
	Application("VERSION") = objParam.GetValue("VERSION")

	Set objParam = Nothing
	
	
	' Load Langauge Label & Strings
	
	
End Sub


' --------------------------------------------------------------------------
' 
' Routine: Sub CheckAuthentication
'
' Purpose: Caches all language labels and associated texts for speedy use
'		   on the asp pages.  The language to be cached is passed into the
'		   function via the "intLangID" variable.
'
' --------------------------------------------------------------------------
Sub CheckAuthentication

	Dim strRedirectURL
	
		
	If Session(lhd_UserID) = 0 Then
	
	  ' Need to re-authenicate, and after doing so redirect the user back to the page they were on
	
	  strRedirectURL = "Logon.asp?URL=" & Request.ServerVariables("PATH_INFO")
	  If Len(Request.ServerVariables("QUERY_STRING")) > 0 Then
			reAddr = reAddr & "?" & Request.ServerVariables("QUERY_STRING")
		End If
  	Response.Redirect strRedirectURL
  	
	Else
	
	  ' No need to re-authenicate so do nothing
	
	End If
	
End Sub


Public Function GetUserID()

	Dim strRedirectURL


	If Session("lhd_UserID") > 0 Then

		GetUserID = Session("lhd_UserID")

	Else

		If Len(Request.ServerVariables("QUERY_STRING")) > 0 Then
			strRedirectURL = "Logon.asp?URL=http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("PATH_INFO") & "?" & Request.ServerVariables("QUERY_STRING")
		Else
			strRedirectURL = "Logon.asp?URL=http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("PATH_INFO")
		End If

		Response.Redirect strRedirectURL
	  
	End If

End Function


Public Function GetUserPermMask()

	If Session("lhd_UserPermMask") > 0 Then
		GetUserPermMask = Session("lhd_UserPermMask")
	Else
		DisplayError 3, "Session timed out.  Click <A href=""Logon.asp"">here</A> to logon."
	End If

End Function


' Loads the standardised error page

Sub DisplayError(ByVal intType, ByVal strErrorDesc)

	Dim strHTML


	' Create the web page
	Response.Clear
	
	strHTML = "<HTML>" &_
			  "<LINK rel=""stylesheet"" type=""text/css"" href=""Default.css"">" &_
			  "<BODY><CENTER><TABLE class=normal width=680  border=0 cellSpacing=1 cellPadding=1>" &_
			  "<TR><TD>" & DisplayHeader() & "</TD></TR>" &_
			  "<TR><TD><TABLE class=""lhd_Box"" width=100% border=0 cellSpacing=0 cellPadding=1>"

'	Error Types:1: Missing required field, required field is parse via "strErrorDesc"
'				2: SQL error
'				3: Generic Error, just display full component string
'				4: Permission Denied, Display full component string

	Select Case intType
	
		Case 1
			strHTML = strHTML & "<TR class=""lhd_Heading1""><TD colspan=3 align=center><B>Required Field</B></TD></TR>" &_
								"<TR><TD colspan=3></TD></TR>" &_
								"<TR>" &_
								"<TR><TD width=10%></TD><TD width=80% align=""left"">The field, <B>" & strErrorDesc & "</B>, is a required field.  Please go back and populate this field before continuing.</TD><TD width=10%></TD>" &_
								"</TR>" &_
								"<TR><TD colspan=3></TD></TR>"

		Case 2
			strHTML = strHTML & "<TR class=""lhd_Heading1""><TD colspan=3 align=center><B>SQL Error</B></TD></TR>" &_
								"<TR><TD colspan=3></TD></TR>" &_
								"<TR>" &_
								"<TD width=10%></TD><TD width=80% align=left>SQL query has failed</TD><TD width=10%></TD>" &_
								"</TR>" &_
								"<TR><TD colspan=3></TD></TR>"
			
		Case 3
			strHTML = strHTML & "<TR class=""lhd_Heading1""><TD colspan=3 align=center><B>General Error</B></TD></TR>" &_
								"<TR><TD colspan=3></TD></TR>" &_
								"<TR>" &_
								"<TD width=10%></TD><TD width=80% align=left>" & strErrorDesc & "</TD><TD width=10%></TD>" &_
								"</TR>" &_
								"<TR><TD colspan=3></TD></TR>"
			
		Case 4
			strHTML = strHTML & "<TR class=""lhd_Heading1""><TD colspan=3 align=center><B>Access Denied</B></TD></TR>" &_
								"<TR><TD colspan=3></TD></TR>" &_
								"<TR>" &_
								"<TD width=10%></TD><TD width=80% align=left>You do not have permission to access this page.  Please conatct your system administrator for further information.</TD><TD width=10%></TD>" &_
								"</TR>" &_
								"<TR><TD colspan=3></TD></TR>"
		
		Case Else
			' Do nothing
			
	End Select

	strHTML = strHTML & "</TABLE></TD></TR>" &_
						"<TR><TD>" & DisplayFooter() & "</TD></TR>" &_
						"</CENTER></BODY></HTML>"


	' Display the page
	Response.Write strHTML

	' Stop processing the .asp file
	Response.End()

End Sub


' Loads the standardised page header

Function DisplayHeader()

	Dim strHTML
	
	strHTML = "<TABLE class=""lhd_Table_Normal"" width=""100%"" cellspacing=""0"">"
	strHTML = strHTML & "<TR style=""Font-Size: 10pt"">"
	strHTML = strHTML & "<TD width=""50%"" style=""BORDER-BOTTOM: solid 2px"" align=""Left"" valign=""Top"">&nbsp;</TD>"
	strHTML = strHTML & "<TD width=""50%"" style=""BORDER-BOTTOM: solid 2px"" align=""Right"" valign=""Top"">" & Application("SITE_NAME") & "</TD>"
	strHTML = strHTML & "</TR>"
	strHTML = strHTML & "</TABLE>"

	DisplayHeader = strHTML

End Function


' Loads the standardised page footer

Function DisplayFooter()

	Dim strHTML
	
	strHTML = "<TABLE class=normal width=100%>"
	strHTML = strHTML & "<BR>"
	strHTML = strHTML & "<TR style=""Font-Size: 8pt""><HR>"
	strHTML = strHTML & "<TD aligh=left width=25%>" & Lang("User_Name") & ": " & Session("lhd_Username") & "</TD>"
	strHTML = strHTML & "<TD align=right width=75%>"
	strHTML = strHTML & "<A href=""Menu.asp"">" & Lang("Main_Menu") & "</A>"

	If PERM_ACCESS_ADMIN = (PERM_ACCESS_ADMIN And Session("lhd_UserPermMask")) Then
         strHTML = strHTML & "&nbsp;&nbsp;|&nbsp;&nbsp;<A href=""admMenu.asp"">" & Lang("Manage_&_Configure_System") & "</A>&nbsp;"
    Else
         ' Do nothing				        
    End If

	strHTML = strHTML & "&nbsp;&nbsp;|&nbsp;&nbsp;<A href=""caseSearch.asp"">" & Lang("Search") & "</A>"
	strHTML = strHTML & "&nbsp;&nbsp;|&nbsp;&nbsp;<A href=""Logoff.asp"">" & Lang("Log_Off") & "</A>"
	strHTML = strHTML & "</TD>"
	strHTML = strHTML & "</TR>"
	strHTML = strHTML & "<TR style=""Font-Size: 8pt"">"
	strHTML = strHTML & "<TD colspan=2>"
	strHTML = strHTML & "Transworld Interactive Help Desk, Copyright (C) 2007&nbsp;&nbsp;&nbsp;( <A href=""License.htm"">" & Lang("View_License") & "</A> )"
	strHTML = strHTML & "</TD>"
	strHTML = strHTML & "</TR>"
	strHTML = strHTML & "</TABLE>"
	
	DisplayFooter = strHTML

End Function


' Loads the the standard page numbers

Function DisplayPageNumbers(  ByVal strPage, ByVal intPage, ByVal intPages )

	Dim strHTML
	Dim I


	If intPages > 1 Then
		strHTML = strHTML & "<TR style=""FONT-WEIGHT: Bold"">"
		strHTML = strHTML & "<TD align=center style=""" & "FONT-SIZE: 9.5pt" & """>"

		If intPage > 1 Then
			strHTML = strHTML & "<A HREF=""" &  strPage & "Page=" & CStr(intPage-1) & """>" & Lang("Previous") & "</A>"
		Else
			strHTML = strHTML & "<FONT color=""gray"">" & Lang("Previous") & "</FONT>"
		End If

		strHTML = strHTML & "&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;"

		If intPages > intPage Then
			strHTML = strHTML & "<A HREF=""" & strPage & "Page=" & CStr(intPage+1) & """>" & Lang("Next") & "</A>"
		Else
			strHTML = strHTML & "<FONT color=""gray"">" & Lang("Next") & "</FONT>"
		End If

		strHTML = strHTML & "</TD></TR>"

	Else
				
		' Do nothing
					
	End If


	DisplayPageNumbers = strHTML


End Function


Sub ValidateContact(ByVal ContactID, ByVal binRequiredPerm)

	Dim objContact
	Dim strURL, strRedirectURL
	
	
	If ContactID = 0 Then
	
		' Need to re-validate the contact

		strRedirectURL = "Logon.asp?URL=" & Request.ServerVariables("PATH_INFO")
		
		If Len(Request.ServerVariables("QUERY_STRING")) > 0 Then
		
			strURL = strRedirectURL & "?" & Request.ServerVariables("QUERY_STRING")
		
		Else
		
			uURL = strRedirectURL
				
		End If
		
		Response.Clear
		Response.Redirect strURL
		
	Else
	
		' Verify the contact is still valid and active
		
		Set objContact = New clsContact
	
		objContact.ID = ContactID	
	
		If Not objContact.Load Then
		
			Response.Write objContact.LastError
			
		Else
		
			If objContact.IsActive Then
			
				' Contact is active now verify they have the required permissions
				
				If objContact.Role.RoleMask And binRequiredPerm Then
				
					' Access granted
					
				Else
				
					' Access denied, need to raise error
					
					Call DisplayError(3, "Access denied.  You do not have permissions required to view this page.")
				
				End If
			
			Else
			
				' Access denied, Contact has been deactivated, need to raise an error
			
				Call DisplayError(3, "Access denied.  You have been deactivated.")
					
			End If
	
		End If
		
		Set objContact = Nothing
		
	End If

End Sub


Function ParseEMail( ByVal strText, ByVal lngCaseID )

	Dim objCase
	
	
	If Len(strText) > 0 Then
		
		Set objCase = New clsCase
	
		objCase.ID = Clng(lngCaseID)
	
		If Not objCase.Load Then
	
			objCase.LastError
	
		Else
	
			strText = Replace(strText, "[CASEID]", objCase.ID)
			strText = Replace( strText, "[TITLE]", objCase.Title)
			strText = Replace(strText, "[DESCRIPTION]", objCase.Description)
			strText = Replace(strText, "[STATUS]", objCase.Status.ItemName)
'			strText = Replace(strText, "[RaisedDate]", DisplayDate(objCase.RaisedDate))
'			strText = Replace(strText, "[ClosedDate]", DisplayDate(objCase.ClosedDate))
'			strText = Replace(strText, "[CATEGORY]", objCase.Cat.CatName)
'			strText = Replace(strText, "[DEPARTMENT]", objCase.Dept.DeptName)
'			strText = Replace(strText, "[phone]", objCase.Contact.OfficePhone)
'			strText = Replace(strText, "[location]", objCase.Contact.Location)
			strText = Replace(strText, "[RESOLUTION]", objCase.Resolution)
			strText = Replace(strText, "[BASE_URL]", Application("BASE_URL"))
			strText = Replace(strText, "[CONTACT_USERNAME]", objCase.Contact.UserName)
			strText = Replace(strText, "[CONTACT_FULLNAME]", objCase.Contact.FName & " " & objCase.Contact.LName)
			strText = Replace(strText, "[CONTACT_EMAIL]", objCase.Contact.Email)
'			strText = Replace(strText, "[repusername]", objCase.Rep.UserName)
'			strText = Replace(strText, "[repfullname]", objCase.Rep.FName & " " & objCase.Rep.LName)
'			strText = Replace(strText, "[repemail]", objCase.Rep.Email)
			strText = Replace(strText, "[URL]", Application("BASE_URL") & "/caseModify.asp?ID=" & objCase.ID)

		End If

		Set objCase = Nothing

	Else
		
		' Do nothing
			
	End If

	' Return the parsed string
	
	ParseEMail = strText
	

End Function


' Sends mail

Sub SendEMail( ByVal lngCaseID, ByVal strEMailMsgType, ByVal strRecipientType )

	Dim objMail, objEMailMsg, objCase, objContact
	Dim I
	Dim strRecipient, strCc
	

	Set objEMailMsg = New clsEMailMsg
	
	objEMailMsg.EMailMsgType = strEMailMsgType
	
	If Not objEMailMsg.Load Then
	
		' Display Error EMailMsgType not returned
			
	Else
	
		Set objMail = New clsMail
		
		objMail.MailMethod = Application("EMAIL_METHOD")
		
		If Len(strRecipientType) > 0 Then
		
			Set objCase = New clsCase
			
			objCase.ID = CLng(lngCaseID)
			
			If Not objCase.Load Then
			
				' Raise Error, case didn't load
				
			Else
		
				Select Case strRecipientType
		
					Case "USER"
						' Send to the Contact and all users listed in the Cc field
						If Len(objCase.AltEmail) > 0 Then
  							objMail.AddRecipient objCase.AltEmail
						Else
  							objMail.AddRecipient objCase.Contact.EMail
						End If
						
						If Len(objCase.Cc) > 0 Then
						
						  strCc = Replace(objCase.Cc, ",", ";")

						  Do
						    
						    I = Instr(1, strCc, ";")
						    
						    If I = 0 Then

								  strRecipient = strCc
								  strCc = ""

    						Else
    						
								  strRecipient = Mid(strCc, 1, I-1)

								  If Instr(1, strRecipient, "@") > 0 Then

    							  objMail.AddRecipient strRecipient

    							Else
    						
    							  Set objContact = New clsContact
    							  objContact.UserName = strRecipient
    							  If Not objContact.Load Then
    							    ' Contact not found
    							  Else
      								objMail.AddRecipient objContact.Email
    							  End If
    							  Set objContact = Nothing
    						
    							End If

  								strCc = Trim(Mid(strCc, I + 1, Len(strCc) - I))

    						End If

						  Loop Until Len(strCc) = 0
						  
						Else
						
						  ' Do nothing
						  
						End If
	
					Case "REP"
						' Send to the Rep
						objMail.AddRecipient objCase.Rep.EMail
						
					Case "GROUP"
						' Send to all members in the RepGroup assigned to the Case
						If objCase.Group.GroupMembers.BOF And objCase.Group.GroupMembers.EOF Then
						
							' No members of the group
						
						Else
						
							' Send an email to all members of the assigned group
							While Not objCase.Group.GroupMembers.EOF
							
								objMail.AddRecipient objCase.Group.GroupMembers.Item.EMail
								objCase.Group.GroupMembers.MoveNext
							
							WEnd
						
						End If
					
					Case "ADMIN"
						' Send an email to the system administrators address
						objMail.AddRecipient Application("SYSTEM_EMAIL")
					
					Case Else
						' Raise Error
						
						' We could use this to send to the alternate email address, which
						' would be passed in via the strRecipientType variable
						objMail.AddRecipient strRecipientType
		
				End Select
				
			End If
			
			Set objCase = Nothing
			
		Else
		
			' If all else fails send an email to the system administrators address
			objMail.AddRecipient Application("SYSTEM_EMAIL")
		
		End If

		
		objMail.From = Application("SYSTEM_EMAIL")
		objMail.Subject = ParseEMail( objEMailMsg.Subject, lngCaseID )
		objMail.Body = ParseEMail( objEMailMsg.Body, lngCaseID )

		objMail.Send
		
		Set objMail = Nothing

	End If
	
	Set objEMailMsg = Nothing

End Sub


' This function is used to build a select list based on the data contained in the tblLists table

Function BuildList( ByVal strListName, ByVal intListItemSelected )

	Dim strHTML
	Dim objCollection
	

	Set objCollection = New clsCollection
					    
	objCollection.CollectionType = objCollection.clListItem
	objCollection.Query = "SELECT * FROM tblLists WHERE ParentListItemFK=(SELECT ListItemPK FROM tblLists WHERE ItemName='"&  strListName & "') AND IsActive=" & lhd_True & " ORDER BY ItemOrder ASC"
					    
	If Not objCollection.Load Then
					    
		' Didn't load
							
	Else
					    
	    If objCollection.BOF And objCollection.EOF Then
					    
			' No records returned
								
		Else
		
			strHTML = ""
							
			Do While Not objCollection.EOF
	 			If objCollection.Item.ID = intListItemSelected Then
					strHTML = strHTML & "<OPTION SELECTED VALUE=""" & objCollection.Item.ID & """>" & objCollection.Item.ItemName & "</OPTION>"
				Else
					strHTML = strHTML & "<OPTION VALUE=""" & objCollection.Item.ID & """>" & objCollection.Item.ItemName & "</OPTION>"
				End If
									
				objCollection.MoveNext
			Loop
							
		End If
					    
	End If

	Set objCollection = Nothing

	
	' Return the select list
	
	BuildList = strHTML


End Function


Function Lang(ByVal strLabel)

	Dim objContact
	Dim intLangID, intUserID
	Dim strCachedLabel, strSQL
	Dim rstLangText

	
	intUserID = Session("lhd_UserID")

	If IsNull(Session("lhd_UserID")) Or IsEmpty(Session("lhd_UserID")) Or Len(Session("lhd_UserID")) = 0 Then
	
		intLangID = Application("DEFAULT_LANGUAGE")
		
	Else
	
		intLangID = Session("lhd_LangID")
		
		If intLangID > 0 Then
		
			' Do nothing as a language has been set
		
		Else
			
			Set objContact = New clsContact
			
			objContact.ID = intUserID
			
			If Not objContact.Load Then
			
				' Raise Error, contact didn't load
			
			Else
	
				If IsNull(objContact.LangID) Or IsEmpty(objContact.LangID) Or objContact.LangID = 0 Then
					intLangID = Application("DEFAULT_LANGUAGE")
				Else
					intLangID = objContact.LangID
				End If
	
				Session("lhd_LangID") = intLangID
				
			End If
			
			Set objContact = Nothing

		End If
		
	End If


	strCachedLabel = Application("lhd_" & strLabel)
	
	If IsEmpty(strCachedLabel) Then
	
		' Need to reload the Language label and associate text from the database
	
		strSQL = "SELECT LangText FROM tblLanguageTexts " &_
				     "WHERE LangLabelFK=(SELECT LangLabelPK FROM tblLanguageLabels WHERE LangLabel='" & strLabel & "') " &_
						 "AND LangFK=" & intLangID

		Set rstLangText = Server.CreateObject("ADODB.Recordset")

		rstLangText.Open strSQL, cnnDB
		
		If rstLangText.BOF And rstLangText.EOF Then
		
			' No records returned, missing label and associated text
			Application("lhd_" & strLabel) = "@" & strLabel & "@"
			
'			' -------------------------------------------------------------------------------------
'			'
'			' Let Temporarily inset new labels into the tblLanguageLabels and then tblLanguageTexts
'
'			rstLangText.Close 
'			Set rstLangText = Nothing
'
'			Dim lnglangLabelID
'			Dim rstLangLabel
'
'			Set rstLangLabel = Server.CreateObject("ADODB.Recordset")
'
'			rstLangLabel.Open "tblLanguageLabels", cnnDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect
'			rstLangLabel.AddNew
'			rstLangLabel.Fields("LangLabel") = strLabel
'			rstLangLabel.Update
'
'			lngLangLabelID = rstLangLabel.Fields("LangLabelPK")
'
'			rstLangLabel.Close
'
'			Set rstLangLabel = Nothing
'
'
'			Set rstLangText = Server.CreateObject("ADODB.Recordset")
'
'			rstLangText.Open "tblLanguageTexts", cnnDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect
'			rstLangText.AddNew
'			rstLangText.Fields("LangFK") = Application("DEFAULT_LANGUAGE")
'			rstLangText.Fields("LangLabelFK") = lngLangLabelID
'			rstLangText.Fields("LangText") = Replace(strLabel, "_", " ")
'			rstLangText.Update
'
'			Application("lhd_" & strLabel) = rstLangText.Fields("LangText")
'
'			' -------------------------------------------------------------------------------------
		  
		Else
		
			Application("lhd_" & strLabel) = rstLangText.Fields("LangText")
		
		End If
		
		rstLangText.Close 
		Set rstLangText = Nothing
		
		strCachedLabel = Application("lhd_" & strLabel)

	Else
	
		' Do nothing
	
	End If
	

'	Lang = "#" & strCachedLabel
	Lang = strCachedLabel
	
	
End Function


' Converts one time format into another.  This function expects a date or time formatted string

Function FormatTime( ByVal dtTime, ByVal strFormat)

  If Not IsDate(dtTime) Then

    FormatTime = ""
    
  Else
    
    Select Case strFormat
  
      Case "hh:mm"
        FormatTime = Right("0" & Hour(dtTime), 2) & ":" & Right("0" & Minute(dtTime), 2)
          
      Case Else
        FormatTime = Hour(dtTime) & ":" & Minute(dtTime)
        
    End Select

  End If

End Function


' Converts one date format into another.  This function expects a date formatted string

Function FormatDate( ByVal dtDate, ByVal strFormat)

  Dim strDay, strMonth, strYear
  

  If Not IsDate(dtDate) Then
  
    ' Raise error as passed string is not in correct format
    FormatDate = ""
  
  Else
    
    strDay = Right("0" & Day(CDate(dtDate)), 2)
    strMonth = Right("0" & Month(CDate(dtDate)), 2)
    strYear = Right("20" & Year(CDate(dtDate)), 4)
    
    Select Case strFormat
  
      Case "dd/mm/yyyy"
        FormatDate = strDay & "/" & strMonth & "/" & strYear
  
      Case "mm/dd/yyyy"
        FormatDate = strMonth & "/" & strDay & "/" & strYear
      
      Case "dd-mmm-yyyy"
        Select Case strMonth

          Case "01"
            strMonth = "Jan"
          Case "02"
            strMonth = "Feb"
          Case "03"
            strMonth = "Mar"
          Case "04"
            strMonth = "Apr"
          Case "05"
            strMonth = "May"
          Case "06"
            strMonth = "Jun"
          Case "07"
            strMonth = "Jul"
          Case "08"
            strMonth = "Aug"
          Case "09"
            strMonth = "Sep"
          Case "10"
            strMonth = "Oct"
          Case "11"
            strMonth = "Nov"
          Case "12"
            strMonth = "Dec"
          Case Else
            ' Do nothing
        
        End Select
        
        FormatDate = strDay & "-" & strMonth & "-" & strYear

      Case "dd.mm.yyyy"
        FormatDate = strDay & "." & strMonth & "." & strYear
  
      Case "mm.dd.yyyy"
        FormatDate = strMonth & "." & strDay & "." & strYear
      
      Case Else
        FormatDate = strMonth & "/" & strDay & "/" & strYear

    End Select  
      
  End If
 
End Function


' This function expects a date/time formatted string and converts it into the format
' specified in the system configuration

Function DisplayDateTime( ByVal dtDateTime)

  Dim strDate, strTime


  strDate = FormatDate(dtDateTime, Application("DATE_FORMAT"))
  strTime = FormatTime(dtDateTime, Application("TIME_FORMAT"))

  DisplayDateTime = Trim(strDate & " " & strTime)

End Function

'Validate String input
Sub ValidateStr(strInput,lCanBeBlank,strField, intMaxLength, ByRef strTarget )
	If (Not (lCanBeBlank OR lCanBeBlank = 1)) And Trim(strInput) = "" Then
		'Response.Write "lCanBeBlank = '" & lCanBeBlank & "' <br>"
		'Response.End
		Call DisplayError(1,strField)
	End If
	If Len(strInput) > intMaxLength Then
		Call DisplayError(3,strField & " Must be less than " & intMaxLength & " characters long.")
	End If
	strTarget = Replace(RTrim(strInput),"'","''")
End Sub
'Validate Integer input
Sub ValidateInt(strInput,intBadValue,strField,ByRef intTarget)
	If (Not IsNumeric(strInput)) Then
		Call DisplayError(3, strField & " must be numeric.")
	ElseIf CInt(strInput) = intBadValue Then
		Call DisplayError(3, strField & " Cannot be " & intBadValue)
	Else
		intTarget = CInt(strInput)
	End If
End Sub
'Validate Long Integer input
Sub ValidateLng(strInput,intBadValue,strField,ByRef intTarget)
	If (Not IsNumeric(strInput)) Then
		Call DisplayError(3, strField & " must be numeric.")
	ElseIf CInt(strInput) = intBadValue Then
		Call DisplayError(3, strField & " Cannot be " & intBadValue)
	Else
		intTarget = CLng(strInput)
	End If
End Sub
'Validate a True or False value
Sub Validatebln(strInput,strField,ByRef blnTarget)
	If (NOT IsNumeric(strInput)) Then
		Call DisplayError(3, strField & " must be numeric.  Press the Back button")
	ElseIf NOT (Cint(strInput) = 0 OR Cint(strInput) =1)  Then
		Call DisplayError(3, strField & " must be 0 or 1.  Press the Back button")
	Else
		blnTarget = CInt(strInput)
	End If
End Sub


Function ValidateTime(ByVal strHour, ByVal strMinute)

  If IsNumeric(strHour) Then
  
    If CInt(strHour) > 0 And CInt(strHour) < 24 Then
      ValidateTime = True
    Else
      ValidateTime = False
    End If

    If IsNumeric(strMinute) Then
    
      If CInt(strMinute) > 0 And CInt(strMinute) < 60 Then
        ValidateTime = True
      Else
        ValidateTime = False
      End If
      
    Else
    
      ValidateTime = False
    
    End If
  
  Else

    ValidateTime = False
    
  End If

End Function


Function ValidateDate(ByVal strDay, ByVal strMonth, ByVal strYear)

  Dim intDay, intMonth, intYear
  Dim blnIsLeapYear
  Dim blnIsValidDay, blnIsValidMonth, blnIsValidYear


  If IsNumeric(strDay) Then
    intDay = CInt(strDay)
  Else
    blnIsValidDay = False
  End If

  If IsNumeric(strMonth) Then
    intMonth = CInt(strMonth)
  Else
    blnIsValidMonth = False
  End If

  If IsNumeric(strYear) Then
    intYear = CInt(strYear)
  Else
    blnIsValidYear = False
  End If


  If intMonth > 12 Then

    blnIsValidMonth = False

  Else
            
    If intDay = 0 Or intDay > 31 Then
              
      blnIsValidDay = False
                
    Else
              
      Select Case intMonth

        Case 1
          If intDay > 31 Then
            blnIsValidDay = False
          End If
        Case 2
          If intYear Mod 4 = 0 Then
            If intYear Mod 100 = 0 Then
              If  intYear Mod 400 = 0 Then
                blnIsLeapYear = True
              Else
                blnIsLeapYear = False
              End If
            Else
              blnIsLeapYear = True
            End If
          Else
            blnIsLeapYear = False
          End If

          If blnIsLeapYear = True Then
            If intDay > 29 Then
              blnIsValidDay = False
            End If
          Else
            If intDay > 28 Then
              blnIsValidDay = False
            End If
          End If
        Case 3
          If intDay > 31 Then
            blnIsValidDay = False
          End If
        Case 4
          If intDay > 30 Then
            blnIsValidDay = False
          End If
        Case 5
          If intDay > 31 Then
            blnIsValidDay = False
          End If
        Case 6
          If intDay > 30 Then
            blnIsValidDay = False
          End If
        Case 7
          If intDay > 31 Then
            blnIsValidDay = False
          End If
        Case 8
          If intDay > 31 Then
            blnIsValidDay = False
          End If
        Case 9
          If intDay > 30 Then
            blnIsValidDay = False
          End If
        Case 10
          If intDay > 31 Then
            blnIsValidDay = False
          End If
        Case 11
          If intDay > 30 Then
            blnIsValidDay = False
          End If
        Case 12
          If intDay > 31 Then
            blnIsValidDay = False
          End If
        Case Else
          blnIsValidMonth = False
            
      End Select
                
    End If
            
  End If

  ValidateDate = False
  
  If blnIsValidDay Then
		DisplayError 3, "Invalid 'Day' component, please re-check the date you have entered."
  Else
    ValidateDate = True
  End If

  If blnIsValidMonth Then
		DisplayError 3, "Invalid 'Month' component, please re-check the date you have entered."
  Else
    ValidateDate = True
  End If

  If blnIsValidYear Then
		DisplayError 3, "Invalid 'Year' component, please re-check the date you have entered."
  Else
    ValidateDate = True
  End If


End Function


Function SQLDateFormat(ByVal strDay, ByVal strMonth, ByVal strYear)

  Select Case strMonth
      
    Case "1"
      strMonth = "Jan"
    Case "2"
      strMonth = "Feb"
    Case "3"
      strMonth = "Mar"
    Case "4"
      strMonth = "Apr"
    Case "5"
      strMonth = "May"
    Case "6"
      strMonth = "Jun"
    Case "7"
      strMonth = "Jul"
    Case "8"
      strMonth = "Aug"
    Case "9"
      strMonth = "Sep"
    Case "10"
      strMonth = "Oct"
    Case "11"
      strMonth = "Nov"
    Case "12"
      strMonth = "Dec"
    Case Else
      ' Do nothing

  End Select
      
  SQLDateFormat = strDay & "-" & strMonth & "-" & strYear
        
End Function


' The following function converts the entered in date to a format that can be uniformly
' stored in the database

Function SQLDate(ByVal strDateTime)

  Dim strDate, strTime, strDay, strMonth, strYear, strHour, strMinute, strTmp
  Dim I
  Dim blnValidDate, blnValidTime


  If IsDate(strDateTime) Then
  
    ' Validate date

    If InStr(1, strDateTime, " ") > 0 Then
      ' There is a time component
      strDate = Mid(strDateTime, 1, InStr(1, strDateTime, " ")-1)
    Else
      ' There is no time component
      strDate = strDateTime
    End If
    

    ' Check Date Formatted Correctly
    Select Case Application("DATE_FORMAT")
  
      Case "dd/mm/yyyy"
        I = InStr(1, strDate, "/")

        If I > 0 Then

          strDay = Mid(strDate, 1, I-1)
        
          strTmp = Mid(strDate, I+1, Len(strDate)-I)
        
          I = InStr(1, strTmp, "/")
          strMonth = Mid(strTmp, 1, I-1)
        
          strTmp = Mid(strTmp, I+1, Len(strTmp)-I)
        
          strYear = Mid(strTmp, 1, Len(strTmp))
  
        Else
        
					DisplayError 3, "Invalid date format, please use 'dd/mm/yyyy hh:mm' format."
        
        End If

      Case "mm/dd/yyyy"
        I = InStr(1, strDate, "/")

        If I > 0 Then

          strMonth = Mid(strDate, 1, I-1)
        
          strTmp = Mid(strDate, I+1, Len(strDate)-I)
        
          I = InStr(1, strTmp, "/")
          strDay = Mid(strTmp, 1, I-1)
        
          strTmp = Mid(strTmp, I+1, Len(strTmp)-I)
        
          strYear = Mid(strTmp, 1, Len(strTmp))
  
        Else
        
					DisplayError 3, "Invalid date format, please use 'mm/dd/yyyy hh:mm' format."
        
        End If

      Case "dd-mmm-yyyy"
        I = InStr(1, strDate, "-")

        If I > 0 Then
        
          strDay = Mid(strDate, 1, I-1)
        
          strTmp = Mid(strDate, I+1, Len(strDate)-I)
        
          I = InStr(1, strTmp, "-")
          strMonth = Mid(strTmp, 1, I-1)
        
          Select Case UCase(strMonth)
                
            Case "JAN"
              strMonth = "1"
            Case "FEB"
              strMonth = "2"
            Case "MAR"
              strMonth = "3"
            Case "APR"
              strMonth = "4"
            Case "MAY"
              strMonth = "5"
            Case "JUN"
              strMonth = "6"
            Case "JUL"
              strMonth = "7"
            Case "AUG"
              strMonth = "8"
            Case "SEP"
              strMonth = "9"
            Case "OCT"
              strMonth = "10"
            Case "NOV"
              strMonth = "11"
            Case "DEC"
              strMonth = "12"
            Case Else
              strMonth = ""
                
          End Select

          strTmp = Mid(strTmp, I+1, Len(strTmp)-I)
        
          strYear = Mid(strTmp, 1, Len(strTmp))
          
        Else
        
					DisplayError 3, "Invalid date format, please use 'dd-mmm-yyyy hh:mm' format."
        
        End If
      
      Case "dd.mm.yyyy"
        I = InStr(1, strDate, ".")

        If I > 0 Then

          strDay = Mid(strDate, 1, I-1)
        
          strTmp = Mid(strDate, I+1, Len(strDate)-I)
        
          I = InStr(1, strTmp, ".")
          strMonth = Mid(strTmp, 1, I-1)
        
          strTmp = Mid(strTmp, I+1, Len(strTmp)-I)
        
          strYear = Mid(strTmp, 1, Len(strTmp))
  
        Else
        
					DisplayError 3, "Invalid date format, please use 'dd/mm/yyyy hh:mm' format."
        
        End If
      
      Case "mm.dd.yyyy"
        I = InStr(1, strDate, ".")

        If I > 0 Then

          strMonth = Mid(strDate, 1, I-1)
        
          strTmp = Mid(strDate, I+1, Len(strDate)-I)
        
          I = InStr(1, strTmp, ".")
          strDay = Mid(strTmp, 1, I-1)
        
          strTmp = Mid(strTmp, I+1, Len(strTmp)-I)
        
          strYear = Mid(strTmp, 1, Len(strTmp))
  
        Else
        
					DisplayError 3, "Invalid date format, please use 'mm/dd/yyyy hh:mm' format."
        
        End If
      
      Case Else
        ' Use default format

    End Select

    blnValidDate = ValidateDate(strDay, strMonth, strYear)
  
  
    ' Validate the time

    strTime = FormatDateTime(CDate(strDateTime), vbShortTime)

    Select Case  Application("TIME_FORMAT")
  
      Case "hh:mm"
        strHour = Mid(strTime, 1, Instr(1, strTime, ":")-1)
        strMinute = Mid(strTime, Instr(1, strTime, ":")+1, Len(strTime)-Instr(1, strTime, ":"))

      Case Else
        ' Do nothing

    End Select
    
    blnValidTime = ValidateTime(strHour, strMinute)


    If blnValidDate Then
      
      If blnValidTime Then
        SQLDate = SQLDateFormat(strDay, strMonth, strYear) & " " & strHour & ":" & strTime
      Else
        SQLDate = SQLDateFormat(strDay, strMonth, strYear) & " 00:00"
      End If
        
    Else

			DisplayError 3, "Invalid date format, please re-check the date format used."

    End If

  Else
  
    If Len(strDateTime) > 0 Then
			DisplayError 3, "Invalid date format, please re-check the date format used." & strDateTime & "."
    Else
      SQLDate = ""
    End If
    
  End If


End Function

'Check to see if user has permission to modify the case.
Function CanModifyCase(objCase,intUserID,binUserPermMask)

Dim blnCanModify

blnCanModify = False

		If PERM_MODIFY_ALL = (PERM_MODIFY_ALL And binUserPermMask) Then
			
			' GRANTED: user has access to view all cases
			blnCanModify = True
			
		Else
			
			' DENIED: user only has access to their own cases
			
			' Check is this Case belongs to this user.
			
			If (objCase.ContactID = intUserID) Or (InStr(1, objCase.Cc, Session("Username")) > 0) Then

				If PERM_MODIFY_OWN = (PERM_MODIFY_OWN And binUserPermMask) Then
				
					' GRANTED: user has access to view this case
					blnCanModify = True
				
				Else
				
					' DENIED: user has no MODIFY access to view this case
				
				End If

			Else
			
				If objCase.RepID = intUserID Then
				
					If PERM_MODIFY_ASSIGNED = (PERM_MODIFY_ASSIGNED And binUserPermMask) Then
					
						' GRANTED: This user is the assiged rep to this case
						blnCanModify = True
						
					Else
	
						' DENIED: The assigned Rep has no MODIFY access this case
					
					End If
				
				Else
				
					' Finally we need to check if the person logged in belongs to the
					' assigned RepGroupPK.
					
					If objCase.Group.IsMember(intUserID) Then

						If PERM_MODIFY_GROUP = (PERM_MODIFY_GROUP And binUserPermMask) Then
					
							' GRANTED: This user is the assiged rep to this case
							blnCanModify = True
							
						Else
	
							' DENIED: The assigned Rep has no MODIFY access this case
					
						End If
					
					Else
					
						' DENIED: user doesn't have access to view this case
								
					End If
					
				End If

			End If
			
		End If
	CanModifyCase = blnCanModify
End Function

</SCRIPT>
