<%@ LANGUAGE="VBScript" %>
<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: kbSearch.asp
'  Date:     $Date: 2004/03/11 06:08:42 $
'  Version:  $Revision: 1.4 $
'  Purpose:  Used to search the knowledgebase via keyword(s)
' ----------------------------------------------------------------------------------
%>

<% 

Option Explicit

%>

<!-- #Include File = "Include/Settings.asp" -->
<!-- #Include File = "Include/Public.asp" -->

<%

' Declare variables

Dim cnnDB
Dim intUserID
Dim binUserPermMask


' Create connection to the database
Set cnnDB = CreateConnection

' Check Session variables
intUserID = GetUserID
binUserPermMask = GetUserPermMask

' Check if the user has rights to view the knowledgebase
If (PERM_KB_READ = (PERM_KB_READ And binUserPermMask)) Or (PERM_KB_MODIFY = (PERM_KB_MODIFY And binUserPermMask)) Or (PERM_KB_CREATE = (PERM_KB_CREATE And binUserPermMask)) Then

Else
	DisplayError 4, ""
End If

%>

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
			<TABLE class="lhd_Box" cellSpacing="0">
			  <FORM name="frmSearch" action="kbList.asp" method="post">
				<TR class="lhd_Heading1">
					<TD colspan="7" align="center"><%=Lang("Knowledgebase")%></TD>
				</TR>
		    <TR>
		      <TD width="8%"></TD>
		      <TD width="15%"></TD>
		      <TD width="10%"></TD>
		      <TD width="34"></TD>
		      <TD width="10%"></TD>
		      <TD width="15%"></TD>
		      <TD width="8%"></TD>
		    </TR>
				</TR>
				<TR>
					<TD></TD>
					<TD><%=Lang("Keywords")%>:</TD>
					<TD colspan="3">
					  <INPUT type="text" id="tbxKeywords" name="tbxKeywords" style="WIDTH: 100%;">
					</TD>
					<TD>
					  &nbsp;&nbsp;&nbsp;<INPUT style="WIDTH: 80px; BACKGROUND-COLOR: white" id="btnSearch" name="btnSearch" type="submit" value="<%=Lang("Search")%>">
					</TD>
					<TD></TD>
				</TR>
				<TR>
					<TD></TD>
					<TD></TD>
					<TD colspan="3" style="FONT-SIZE: 8pt;">
					  <BR>
					  <B><%=Lang("Note")%>:</B> A knowledgebase search will search on Reference ID, Issue, Cause and Resolution text.
					</TD>
					<TD></TD>
					<TD></TD>
				</TR>
				<TR>
					<TD colspan="7"></TD>
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
