<%@ LANGUAGE="VBScript" %>
<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: admLanguageList.asp
'  Date:     $Date: 2004/03/17 00:08:28 $
'  Version:  $Revision: 1.7 $
'  Purpose:  Administration page for listing all Languages
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
	

	' Build the list of departments
	
	Set objCollection = New clsCollection
	
	objCollection.CollectionType = objCollection.clLanguage
	objCollection.Query = "SELECT * FROM tblLanguages ORDER BY LangName ASC"
										
	If Not objCollection.Load Then
				
		Response.Write objCollection.LastError
					
	Else
	
		If objCollection.BOF And objCollection.EOF Then
		
			' No records
			
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
				strHTML = strHTML & "	<TD class=LeftColumn  align=left>&nbsp;" & objCollection.Item.LangName & "</TD>" & Chr(13)
				strHTML = strHTML & "	<TD class=LeftColumn  align=left>&nbsp;" & objCollection.Item.Localized & "</TD>" & Chr(13)
				If objCollection.Item.IsActive Then
					strHTML = strHTML & "	<TD class=LeftColumn  align=center>" & Lang("Yes") & "</TD>" & Chr(13)
				Else
					strHTML = strHTML & "	<TD class=LeftColumn  align=center>" & Lang("No") & "</TD>" & Chr(13)
				End If
				strHTML = strHTML & "	<TD class=RightColumn  align=center><A href=""admLanguage.asp?ID=" & objCollection.Item.ID & """><IMG src=""Images/Pencil.gif"" alt=""Edit"" border=""0""></A>&nbsp;&nbsp;&nbsp;&nbsp;<A href=""admLanguageStringList.asp?Language=" & objCollection.Item.ID & "&Page=1""><IMG src=""Images/Manage.gif"" alt=""Manage Strings"" border=""0""></A></TD>" & Chr(13)
				strHTML = strHTML & "</TR>" & Chr(13)
				
				objCollection.MoveNext

			Loop
			
		End If
	
	End If
	
	
	Set objCollection = Nothing

%>

<HTML>
<HEAD>

<META content="MSHTML 6.00.2600.0" name=GENERATOR></HEAD>

<LINK rel="stylesheet" type="text/css" href="Default.css">

<BODY>
<P align=center>
<TABLE align=center cellSpacing=1 cellPadding=1 width="680" border=0>
  
  <TR>
    <TD>
    <%
    Response.Write DisplayHeader
    %>
    </TD>
  </TR>
  <TR>
    <TD>
		<TABLE class=Normal cellSpacing=1 cellPadding=1 width="100%" border=0 style="WIDTH: 100%" bgColor=white>
			<TR class="lhd_Heading1">
				<TD colspan=5 align=center><%=Lang("Languages")%></TD>
   			</TR>
		    <TR class=ColumnHeading>
				<TD class=LeftColumn align=left width="35%">&nbsp;<%=Lang("Language")%></TD>
				<TD class=LeftColumn align=left width="35%">&nbsp;<%=Lang("Localised")%></TD>
				<TD class=LeftColumn align=center width="15%"><%=Lang("Active")%></TD>
				<TD class=RightColumn align=center width="15%"><%=Lang("Options")%></TD>
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
			Response.Write DisplayPageNumbers(  "admLanguageList.asp?", intPage, intPages )
			%>
		    <TR>
				<TD align=right>
					<BR><INPUT style="WIDTH: 80px; BACKGROUND-COLOR: white" OnClick="VBScript:window.navigate 'admLanguage.asp'" id=btnNew name=btnNew type=submit value="<%=Lang("New")%>">
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
