<%@ LANGUAGE="VBScript" %>
<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: admAssignmentList.asp
'  Date:     $Date: 2004/03/11 06:08:42 $
'  Version:  $Revision: 1.6 $
'  Purpose:  Administration page for listing all Assignments
' ----------------------------------------------------------------------------------
%>
<%

Option Explicit



%>

<!--METADATA TYPE="TypeLib" NAME="Microsoft ActiveX Data Objects 2.5 Library" UUID="{00000205-0000-0010-8000-00AA006D2EA4}" VERSION="2.5"-->
<!--METADATA TYPE="TypeLib" NAME="Microsoft Scripting Runtime" UUID="{420B2830-E718-11CF-893D-00A0C9054228}" VERSION="1.0"-->
<!--METADATA TYPE="TypeLib" NAME="Microsoft CDO for Windows 2000 Library" UUID="{CD000000-8B95-11D1-82DB-00C04FB1625D}" VERSION="1.0"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<!-- #Include File = "Include/Settings.asp" -->
<!-- #Include File = "Include/Public.asp" -->

<!-- #Include File = "Classes/clsAssignment.asp" -->
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


<HTML>

<%

	Dim cnnDB
	Dim binUserPermMask, binRequiredPerm
	Dim objCollection
	Dim I, intPages, intPage, intUserID
	Dim strHTML


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


	
	' Get the settings from the QueryString
	
	intPage = CInt(Request.Querystring("Page"))
	If intPage = 0 Then
		intPage = 1
	Else
		' Do nothing
	End If


	' Build the list of assignments
	
	Set objCollection = New clsCollection
	
	objCollection.CollectionType = objCollection.clAssignment
	objCollection.Query = "SELECT * FROM tblAssignments"

	If Not objCollection.Load Then
				
		' Raise Error, Recordset failed to load
					
	Else
										
		If objCollection.BOF And objCollection.EOF Then
					
			' No records returned

		Else
			
			If objCollection.RecordCount Mod Application("ITEMS_PER_PAGE") = 0 Then
				intPages = Int(objCollection.RecordCount / Application("ITEMS_PER_PAGE"))
			Else
				intPages = Int(objCollection.RecordCount / Application("ITEMS_PER_PAGE")) + 1
			End If

			strHTML = ""
			
			' Move the the record at the start of the next page
			
			objCollection.Move(Application("ITEMS_PER_PAGE") * (intPage - 1))

			I = 0
	
			Do While Not objCollection.EOF And Application("ITEMS_PER_PAGE") > I

				I = I + 1
					
				strHTML = strHTML & "<TR>" & Chr(13)
				strHTML = strHTML & "	<TD class=LeftColumn align=left>&nbsp;" & objCollection.Item.CaseType.CaseTypeName & "</TD>" & Chr(13)
				strHTML = strHTML & "	<TD class=LeftColumn align=left>&nbsp;" & objCollection.Item.Cat.CatName & "</TD>" & Chr(13)
				strHTML = strHTML & "	<TD class=LeftColumn align=center>" & objCollection.Item.Rep.Username & "</TD>" & Chr(13)
				If objCollection.Item.IsActive Then
					strHTML = strHTML & "	<TD class=LeftColumn align=center>" & Lang("Yes") & "</TD>" & Chr(13)
				Else
					strHTML = strHTML & "	<TD class=LeftColumn align=center>" & Lang("No") & "</TD>" & Chr(13)
				End If
				strHTML = strHTML & "	<TD class=RightColumn align=center><A href=""admAssignment.asp?ID=" & objCollection.Item.ID & """><IMG src=""Images/Pencil.gif"" alt=""Edit"" border=""0""></A></TD>" & Chr(13)
				strHTML = strHTML & "</TR>" & Chr(13)
					
				objCollection.MoveNext

			Loop
				
		End If
		
	End If
	
	Set objCollection = Nothing

%>

<HEAD>

<META content="MSHTML 6.00.2600.0" name=GENERATOR></HEAD>

<LINK rel="stylesheet" type="text/css" href="Default.css">

<BODY>
<P align=center>
<TABLE align=center cellSpacing=1 cellPadding=0 width="680" border=0>
  
  <TR>
    <TD>
    <%
    Response.Write DisplayHeader
    %>
    </TD>
  </TR>
  <TR>
    <TD>
		<TABLE class=Normal cellSpacing=1 cellPadding=1 width="100%" border=0 bgColor=white>
			<TR class="lhd_Heading1">
				<TD colspan=5 align=center><%=Lang("Assignments")%></TD>
   			</TR>
		    <TR class=ColumnHeading>
		    <TD class=LeftColumn align=left width="25%">&nbsp;<%=Lang("Case_Type")%></TD>
		    <TD class=LeftColumn align=left width="20%">&nbsp;<%=Lang("Category")%></TD>
		    <TD class=LeftColumn align=center width="20%"><%=Lang("Rep")%></TD>
		    <TD class=LeftColumn align=center width="10%"><%=Lang("Active")%></TD>
		    <TD class=RightColumn align=center width="15%"><%=Lang("Actions")%></TD>
		    </TR>
		    <%
		    Response.Write strHTML
		    %>
		</TABLE>
	</TD>
</TR>
<TR>
    <TD>
		<TABLE class=Normal cellSpacing=1 cellPadding=1 width="100%" border=0 bgColor=white>
			<%
			Response.Write DisplayPageNumbers(  "admAssignmentList.asp?", intPage, intPages )
			%>
		    <TR>
				<TD align=right>
					<BR><INPUT style="WIDTH: 80px; BACKGROUND-COLOR: white" OnClick="VBScript:window.navigate 'admAssignment.asp'" id=btnNew name=btnNew type=submit value="<%=Lang("New")%>">
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
	cnnDB.Close
	Set cnnDB = Nothing
%>