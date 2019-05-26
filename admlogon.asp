<% 
Option Explicit

' Buffer the response, so Response.Expires can be used

Response.Buffer = True
Response.Expires = -1
%>
<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: admLogon.asp
'  Date:     $Date: 2004/03/10 08:21:12 $
'  Version:  $Revision: 1.6 $
'  Purpose:  Administration logon page
' ----------------------------------------------------------------------------------
%>


<!-- #Include File = "Include/Public.asp"	-->
<!-- #Include File = "Include/Settings.asp" -->

<!-- #Include File = "Classes/clsContact.asp" -->
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
	  
	    ' Password is valid, so we now need to check they have permission to administer
	    
	    If PERM_ACCESS_ADMIN = (PERM_ACCESS_ADMIN And objContact.Role.RoleMask) Then
      	' Admin access granted

      Else
      	' Admin access denied, Display denied message
      	DisplayError 4, ""
      End If
			  		
	  	Session("lhd_UserID") = objContact.ID
	  	Session("lhd_UserName") = objContact.UserName
	  	Session("lhd_UserPermMask") = objContact.Role.RoleMask
	  	Session("lhd_LangID") = objContact.LangID

			Response.Redirect "admMenu.asp"

  	Else
					
		  ' Invalid Password, the user needs to try again.
		  strLogonResult = "Logon failed, Invalid password."
					
	  End If
					
  End If
      
	Set objContact = Nothing
			
Else
    
	' Do nothing and load the login screen
    
End If  
	
%>

<HTML>

<HEAD>
	
	<META content="MstrHTML 6.00.2600.0" name=GENERATOR>
</HEAD>

<LINK rel="stylesheet" type="text/css" href="Default.css">

<BODY>
<P align=center>
<TABLE class=Normal>
	<TR>
		<TD>
			<%
			Response.Write DisplayHeader
			%>
		</TD>
	</TR>
	<TR>
		<TD>
			<TABLE class="lhd_Box" cellspacing=0>
				<FORM action="admLogon.asp" method="post" id=frmLogon name=frmLogon>
				<INPUT id=tbxLogon name=tbxLogon type=hidden value="1">
				<TR class="lhd_Heading1">
					<TD colspan=5 align=middle><%=Lang("Administration_Log_On")%></TD>
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
				  <TD colspan="5"></TD>
				</TR>
				</FORM>
			</TABLE>
		</TD>
	</TR>
  <TR>
    <TD>
    <%
    Response.Write DisplayFooter
    %>
    </TD>
  </TR>
</TABLE>
</P>
</BODY>
</HTML>

<%

cnnDB.Close
Set cnnDB = Nothing

%>
