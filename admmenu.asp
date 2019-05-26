<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: admMenu.asp
'  Date:     $Date: 2004/03/11 06:08:42 $
'  Version:  $Revision: 1.6 $
'  Purpose:  Administration menu page
' ----------------------------------------------------------------------------------
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>

<META content="MSHTML 6.00.2600.0" name=GENERATOR></HEAD>

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
	Dim cnnDB
	Dim intUserID
	Dim blnUserPermMask, binRequiredPerm


	' Get user variables

    Set cnnDB = CreateConnection
    
    intUserID = GetUserID
    binUserPermMask = GetUserPermMask


	' Check permissions

	If PERM_ACCESS_ADMIN = (PERM_ACCESS_ADMIN And binUserPermMask) Then
		' Admin access granted
		
	Else
		' Here we can either display a denied message or redirect them to the logon screen
		Response.Redirect "admLogon.asp"
	End If

%>

<LINK rel="stylesheet" type="text/css" href="Default.css">

<BODY>
<P align=center>
<TABLE class=Normal align=center cellSpacing=1 cellPadding=1 width="680" border=0>
  
  <TR>
    <TD>
    <%
    Response.Write DisplayHeader
    %>
    </TD>
  </TR>

  <TR>
    <TD>
<TABLE class="lhd_Box" cellSpacing=0 cellPadding=1 width="100%" border=0 bgColor=white>
	<TR class="lhd_Heading1">
		<TD colspan=5 align=center><%=Lang("Administration_Menu")%></TD>
	</TR>
  
  <TR>
    <TD width="20%"></TD>
    <TD width="20%"></TD>
    <TD width="20%"></TD>
    <TD width="20%"></TD>
    <TD width="20%"></TD></TR>
  <TR>
    <TD></TD>
    <TD align=middle colspan=3><A href="admConfiguration.asp"><%=Lang("System_Configuration")%></A></TD>
    <TD></TD></TR>
  <TR>
    <TD></TD>
    <TD align=middle colspan=3><A href="admCaseTypeList.asp?Page=1"><%=Lang("Manage_Case_Types")%></A></TD>
    <TD></TD></TR>
  <TR>
    <TD></TD>
    <TD align=middle colspan=3><A href="admCategoryList.asp?Page=1"><%=Lang("Manage_Categories")%></A></TD>
    <TD></TD></TR>
  <TR>
    <TD></TD>
    <TD align=middle colspan=3><A href="admContactList.asp?Page=1"><%=Lang("Manage_Contacts")%></A></TD>
    <TD></TD></TR>
  <TR>
    <TD></TD>
    <TD align=middle colspan=3><A href="admDepartmentList.asp?Page=1"><%=Lang("Manage_Departments")%></A></TD>
    <TD></TD></TR>
  <TR>
    <TD></TD>
    <TD align=middle colspan=3><A href="admEmailMsgList.asp?Page=1"><%=Lang("Manage_Email_Messages")%></A></TD>
    <TD></TD></TR>
  <TR>
    <TD></TD>
    <TD align=middle colspan=3><A href="admGroupList.asp?Page=1"><%=Lang("Manage_Groups")%></A></TD>
    <TD></TD></TR>
  <TR>
    <TD></TD>
    <TD align=middle colspan=3><A href="admParentListItemList.asp?Page=1"><%=Lang("Manage_Lists")%></TD>
    <TD></TD></TR>
  <TR>
    <TD></TD>
    <TD align=middle colspan=3><A href="admOrganisationList.asp?Page=1"><%=Lang("Manage_Organisations")%></TD>
    <TD></TD></TR>
  <TR>
    <TD> </TD>
    <TD colspan=3 align=middle><A href="admRoleList.asp?Page=1"><%=Lang("Manage_Roles")%></A></TD>
    <TD></TD></TR>
  <TR>
    <TD> </TD>
    <TD colspan=3 align=middle><A href="admLanguageList.asp?Page=1"><%=Lang("Manage_Languages")%></A></TD>
    <TD></TD></TR>
  <TR>
    <TD> </TD>
    <TD colspan=3 align=middle><A href="admAssignmentList.asp?Page=1"><%=Lang("Manage_Assignments")%></A></TD>
    <TD></TD></TR>
  <TR>
    <TD> </TD>
    <TD colspan=3 align=middle></TD>
    <TD></TD>
  </TR>
  </TABLE></TD></TR>
  <TR>
    <TD>
    <%
    Response.Write DisplayFooter
    %>
    </TD>
  </TR>
</TABLE></P>
</BODY></HTML>

<%
cnnDB.Close
Set cnnDB = Nothing
%>