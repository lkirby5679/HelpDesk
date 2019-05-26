<%@ LANGUAGE="VBScript" %>
<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: Logoff.asp
'  Date:     $Date: 2004/03/11 06:08:42 $
'  Version:  $Revision: 1.5 $
'  Purpose:  Used to log a user off
' ----------------------------------------------------------------------------------
%>

<% 
Option Explicit
%>

<!-- #Include File = "Include/Settings.asp" -->
<!-- #Include File = "Include/Public.asp" -->

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

' Declare variables

Dim cnnDB


' Create connection to the database

Set cnnDB = CreateConnection


' Clear the session variables	
	
Session("lhd_UserID") = Empty
Session("lhd_UserName") = Empty
Session("lhd_UserPermMask") = Empty
Session("lhd_LangID") = Empty
	
%>

<HEAD>
	
	<META content="Microsoft FrontPage 6.0" name=GENERATOR>
</HEAD>

<LINK rel="stylesheet" type="text/css" href="Default.css">

<BODY>
<P align=center>
<TABLE class=Normal>
	<TR>
		<TD>
		<a href="http://www.transworldinteractive.net/">
		<img border="0" src="Images/TIisIT.jpg" width="407" height="104"></a><%
		Response.Write DisplayHeader
		%>
		</TD>
	</TR>
	<TR>
		<TD>
			<TABLE class="lhd_Box" cellSpacing=0>
				<TR class="lhd_Heading1">
					<TD colspan=5 align=middle><%=Lang("Log_Off")%></TD>
				</TR>
				<TR>
					<TD width=10%></TD>
					<TD width=20%></TD>
					<TD width=40%></TD>
					<TD width=20%></TD>
					<TD width=10%></TD>
				</TR>
				<TR>
					<TD></TD>
					<TD colspan=3>
						Good bye, your session has been successfully logged out.  ( <A align=center href="Logon.asp"><%=Lang("Logon")%> ...</A> )
					</TD>
					<TD></TD>
				</TR>
				<TR>
					<TD></TD>
					<TD></TD>
					<TD></TD>
					<TD></TD>
					<TD></TD>
				</TR>
			</TABLE>
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