<%@Language="VBScript"%>
<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: kbView.asp
'  Date:     $Date: 2004/03/11 06:08:42 $
'  Version:  $Revision: 1.4 $
'  Purpose:  Provides a read only view of a knowledgebase record
' ----------------------------------------------------------------------------------
%>

<% 

Option Explicit

%>

<HTML>

<!--METADATA TYPE="TypeLib" NAME="Microsoft ActiveX Data Objects 2.5 Library" UUID="{00000205-0000-0010-8000-00AA006D2EA4}" VERSION="2.5"-->
<!--METADATA TYPE="TypeLib" NAME="Microsoft Scripting Runtime" UUID="{420B2830-E718-11CF-893D-00A0C9054228}" VERSION="1.0"-->
<!--METADATA TYPE="TypeLib" NAME="Microsoft CDO for Windows 2000 Library" UUID="{CD000000-8B95-11D1-82DB-00C04FB1625D}" VERSION="1.0"-->

<!-- #Include File = "Include/Settings.asp" -->
<!-- #Include File = "Include/Public.asp" -->

<!-- #Include File = "Classes/clsContact.asp" -->
<!-- #Include File = "Classes/clsKnowledgebase.asp" -->

<%

' Declare variables

Dim cnnDB
Dim intUserID
Dim binUserPermMask
Dim objKnowledgebase

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


Set objKnowledgebase = New clsKnowledgebase

objKnowledgebase.ID = Request.QueryString("ID")

If Not objKnowledgebase.Load Then
  ' Raise error, object failed to load
  
Else
  ' Object loaded

End If


%>

<SCRIPT Language="Javascript">

  function onClick_btnModify(lngID)
  {
//    alert(lngID) ;
    window.navigate("kbRecord.asp?ID=" + lngID) ;
  }

</SCRIPT>


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
			<TABLE class="Normal" cellSpacing=0>
				<TR class="lhd_Heading1">
					<TD colspan=5 align=middle><%=Lang("Knowledgebase_Record")%>&nbsp;&nbsp;(&nbsp;<%="KB" & Right("0000000" & CStr(objKnowledgebase.ID), 8)%>&nbsp;)</TD>
				</TR>
				<TR>
					<TD width="15%"></TD>
					<TD width="20%"></TD>
					<TD width="30%"></TD>
					<TD width="20%"></TD>
					<TD width="15%"></TD>
				</TR>
				<TR class="lhd_Heading2">
					<TD colspan="5"><%=Lang("Issue")%></TD>
				</TR>
				<TR>
					<TD></TD>
					<TD colspan="4"><%=Replace(objKnowledgebase.Issue,  Chr(13) & Chr(10), "<BR>")%></TD>
				</TR>
				<TR>
					<TD colspan="5"><BR><BR></TD>
				</TR>
				<TR class="lhd_Heading2">
					<TD colspan="5"><%=Lang("Cause")%></TD>
				</TR>
				<TR>
					<TD></TD>
					<TD colspan="4"><%=Replace(objKnowledgebase.Cause, Chr(13) & Chr(10), "<BR>")%></TD>
				</TR>
				<TR>
					<TD colspan="5"><BR><BR></TD>
				</TR>
				<TR class="lhd_Heading2">
					<TD colspan="5"><%=Lang("Resolution")%></TD>
				</TR>
				<TR>
					<TD></TD>
					<TD colspan="4"><%=Replace(objKnowledgebase.Resolution,  Chr(13) & Chr(10), "<BR>")%></TD>
				</TR>
				<TR>
					<TD colspan="5"><BR><BR></TD>
				</TR>
				<TR>
					<TD colspan="3" align="Left" style="FONT-SIZE: 8pt;">
					  <BR>
					  <%=Lang("Reference ID")%>: <%="KB" & Right("0000000" & CStr(objKnowledgebase.ID), 8)%>
					  <BR>
					  <%=Lang("Last Updated")%>: <%=DisplayDateTime(objKnowledgebase.LastUpdate)%>
					</TD>
					<TD colspan="2" align="Right" valign="Bottom">
					  <%
		        If PERM_KB_MODIFY = (PERM_KB_MODIFY And binUserPermMask) Then
		        %>
  					  <INPUT style="WIDTH: 80px; BACKGROUND-COLOR: white" id="btnModify" name="btnModify" type="Button" onClick="Javascript: onClick_btnModify(<%=objKnowledgebase.ID%>)" value="<%=Lang("Modify")%>">
  					<%
  					Else
  					  'Do nothing
  					  
  					End If
  					%>
					</TD>
				</TR>
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

Set objKnowledgebase = Nothing

cnnDB.Close
Set cnnDB = Nothing

%>
