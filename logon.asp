<% 
Option Explicit

' Buffer the response, so Response.Expires can be used

Response.Buffer = True
Response.Expires = -1
%>
<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: Logon.asp
'  Date:     $Date: 2004/03/11 06:08:42 $
'  Version:  $Revision: 1.7 $
'  Purpose:  The logon page to allow and authenticate user access
' ----------------------------------------------------------------------------------
%>

<!-- #Include File = "Include/Public.asp" -->
<!-- #Include File = "Include/Settings.asp" -->

<!-- #Include File = "Classes/clsCase.asp" -->
<!-- #Include File = "Classes/clsCaseType.asp" -->
<!-- #Include File = "Classes/clsCategory.asp" -->
<!-- #Include File = "Classes/clsCollection.asp" -->
<!-- #Include File = "Classes/clsContact.asp" -->
<!-- #Include File = "Classes/clsDepartment.asp" -->
<!-- #Include File = "Classes/clsFile.asp" -->
<!-- #Include File = "Classes/clsGroup.asp" -->
<!-- #Include File = "Classes/clsLanguage.asp" -->
<!-- #Include File = "Classes/clsListItem.asp" -->
<!-- #Include File = "Classes/clsMail.asp" -->
<!-- #Include File = "Classes/clsNote.asp" -->
<!-- #Include File = "Classes/clsOrganisation.asp" -->
<!-- #Include File = "Classes/clsParameter.asp" -->
<!-- #Include File = "Classes/clsRole.asp" -->

<%

Dim cnnDB
Dim objContact
Dim strAUTH_USER, strURL, strLogonResult, strPassword, strUsername, strRedirectURL
Dim blnLogon
	            


' Set up the Application Variables

Call SetAppVariables()


' Create connection to the database

Set cnnDB = CreateConnection


' Load all other system defined Parameters into the "Application" variable

Call SetApplicationParams()


'  Cache all the language strings for the default language assigned in the
' configuration table
					
Call CacheLanguageStrings( Application("DEFAULT_LANGUAGE") )



If Len(Trim(Request.QueryString("URL"))) > 0 Then
  strRedirectURL = Trim(Request.QueryString("URL"))
Else
  strRedirectURL = ""
End If

Select Case Application("AUTH_TYPE")
	
	Case lhd_ADAuthentication  ' AD or NT Authentication
		
		strAUTH_USER = Request.ServerVariables("AUTH_USER") 
		
		Set objContact = New clsContact
		
		objContact.UserName = Right(strAUTH_USER, Len(strAUTH_USER) - InStr(strAUTH_USER, "\"))
				
		If Not objContact.Load Then
				
			' Raise Error or need to register the user.

			strURL = "Contact.asp"
					
		Else
		
			' Set the session variables
	            
			Session("lhd_UserID") = objContact.ID
			Session("lhd_UserName") = objContact.UserName
			Session("lhd_UserPermMask") = objContact.Role.RoleMask
			Session("lhd_LangID") = objContact.LangID

			objContact.LogLastAccess


			'  Cache all the language strings for the dusers assigned language
								
			If objContact.LangID <> Application("DEFAULT_LANGUAGE") Then
				Call CacheLanguageStrings( objContact.LangID )
			Else
				' Do nothing
			End If


			strURL = "Menu.asp"

		End If
			
		Set objContact = Nothing


	Case lhd_DBAuthentication  ' DB Authentication

		If Request.Form("tbxLogon") = "1" Then
			blnLogon = True
		Else
			blnLogon = False
		End If
			
		If blnLogon = True Then

      strUsername = Left(Trim(Request.Form("tbxUsername")), 32)
      strPassword = Left(Trim(Request.Form("pwdPassword")), 32)
      	            
      Set objContact = New clsContact
      	            
      objContact.UserName = strUserName
      	            
      If Not objContact.Load Then
      					
        ' Contact not loaded
        strLogonResult = "Logon failed, Invalid username."

      Else
	            
			  ' Validate password

			  If objContact.CheckPassword(strPassword) = True Then
			  		
			  	Session("lhd_UserID") = objContact.ID
			  	Session("lhd_UserName") = objContact.UserName
			  	Session("lhd_UserPermMask") = objContact.Role.RoleMask
			  	Session("lhd_LangID") = objContact.LangID

			  	objContact.LogLastAccess

				  '  Cache all the language strings for the dusers assigned language
				  						
				  If objContact.LangID <> Application("DEFAULT_LANGUAGE") Then
				  	Call CacheLanguageStrings( objContact.LangID )
				  Else
				  	' Do nothing
				  End If

  				strLogonResult = "Logon successful"
					
	  			strURL = "Menu.asp"

		  	Else
					
				  ' Invalid Password, the user needs to try again.
						
				  strLogonResult = "Logon failed, Invalid password."
					
			  End If
					
      End If
      
  		Set objContact = Nothing
			
    Else
    
      ' Do nothing
    
    End If  
	            

	Case Else  ' Catch all other cases

		
End Select

If Len(strURL) > 0 Then

  If Len(strRedirectURL) > 0 Then
	  Response.Redirect strRedirectURL
  Else
	  Response.Redirect strURL
	End If

Else

	' Do nothing and load the login screen

End If
	
%>

<HTML>

<HEAD>
	
	<META content="Microsoft FrontPage 6.0" name=GENERATOR>
<title>Transworld Interactive Help Desk</title>
<script src="../Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
</HEAD>

<LINK rel="stylesheet" type="text/css" href="Default.css">

<BODY>
<TABLE class=Normal>
<TR>
  <TD><script type="text/javascript">
AC_FL_RunContent( 'codebase','http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0','width','778','height','321','title','Transworld Interactive Support','src','support','quality','high','pluginspage','http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash','movie','support' ); //end AC code
</script><noscript><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0" width="778" height="321" title="Transworld Interactive Support">
    <param name="movie" value="support.swf">
    <param name="quality" value="high">
    <embed src="support.swf" quality="high" pluginspage="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="778" height="321"></embed>
  </object></noscript></TD>
</TR>
<TR>
		<TD>
			<p>
			  <%
			Response.Write DisplayHeader
			%>
            </p>
		</TD>
  </TR>
	<TR>
		<TD>
			Please Register if you have not already. You must register before 
			you can enter a problem with any of our services.<TABLE class="lhd_Box" cellspacing=0>
				<FORM action="Logon.asp" method="post" id=frmLogon name=frmLogon>
				<INPUT id=tbxLogon name=tbxLogon type=hidden value="1">
				<TR class="lhd_Heading1">
					<TD colspan=5 align=middle><%=Lang("Log_On")%></TD>
				</TR>
				<TR>
					<TD width=20%></TD>
					<TD width=20%></TD>
					<TD width=20%></TD>
					<TD width=20%></TD>
					<TD width=20%></TD>
				</TR>
				<TR>
					<TD></TD>
					<TD><%=Lang("User_Name")%>:</TD>
					<TD><INPUT id=tbxUsername name=tbxUsername></TD>
					<TD></TD>
					<TD></TD>
				</TR>
				<TR>
					<TD></TD>
					<TD><%=Lang("Password")%>:</TD>
					<TD><INPUT id=pwdPassword type=password name=pwdPassword></TD>
					<TD></TD>
					<TD></TD>
				</TR>
				<TR>
					<TD></TD>
					<TD></TD>
					<TD align=right><INPUT style="BACKGROUND-COLOR: white" id=btnLogon type=submit value="<%=Lang("Log_On")%>" name=btnLogon></TD>
					<TD colspan=2><FONT color=Red>&nbsp;&nbsp;<%=strLogonResult%></FONT></TD>
				</TR>
				<TR>
					<TD></TD>
					<TD></TD>
					<TD style="FONT-SIZE: 8pt" >&nbsp;<%=Lang("New_User")%>? ... <A href="Contact.asp"><%=Lang("Register")%></A></TD>					<TD></TD>
					<TD></TD>
				</TR>
				</FORM>
			</TABLE>		</TD>
	</TR>
</TABLE>
<p></P>
<table width="781" border="0">
  <tr>
    <td width="775" align="center"><p><a href="http://www.transworldinteractive.net/"><img src="Images/TIisIT.jpg" alt="Transworld Interactive Support" width="407" height="104" border="0"></a></p>
    <p><a href="http://www.transworldinteractive.net/"> &copy;&nbsp;1991-2007 Transworld Interactive,Inc. All Rights Reserved.&nbsp;</a></p></td>
  </tr>
</table>
</BODY>
</HTML>

<%

cnnDB.Close
Set cnnDB = Nothing

%>