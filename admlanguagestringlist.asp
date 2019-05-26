<%@ LANGUAGE="VBScript" %>
<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: admLanguageStringList.asp
'  Date:     $Date: 2004/03/18 01:17:25 $
'  Version:  $Revision: 1.4 $
'  Purpose:  Administration page for Listing LanguageStrings
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
<!-- #Include File = "Classes/clsLanguageLabel.asp" -->
<!-- #Include File = "Classes/clsLanguageText.asp" -->
<!-- #Include File = "Classes/clsListItem.asp" -->
<!-- #Include File = "Classes/clsMail.asp" -->
<!-- #Include File = "Classes/clsNote.asp" -->
<!-- #Include File = "Classes/clsOrganisation.asp" -->
<!-- #Include File = "Classes/clsParameter.asp" -->
<!-- #Include File = "Classes/clsRole.asp" -->

<%
	Dim cnnDB
	Dim binUserPermMask, binRequiredPerm
	Dim objCollection, objLanguage
	Dim strTitle
	Dim I, intPages, intPage, intUserID
	Dim strHTML
	Dim intLangID


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
	
	Call ValidateInt(Request.Querystring("Page"),-1,"Page Number",intPage)
	If intPage = 0 Then
		intPage = 1
	Else
		' Do nothing
	End If
	Call ValidateInt(Request.Querystring("Language"),0,"Language ID",intLangID)

	' Build the Language Object
	Set objLanguage = New clsLanguage
	
	objLanguage.ID = intLangID
	objLanguage.Load
	
	strTitle = objLanguage.LangName & "_String_Management"

	' Build the list of Language Texts
	

	Set objCollection = New clsCollection
	
	objCollection.CollectionType = objCollection.clLanguageText
	objCollection.Query = "SELECT * FROM tblLanguageTexts " & _
                        "WHERE LangFK = " & intLangID & " " & _
                        "ORDER BY LangText ASC"

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
				strHTML = strHTML & "	<TD class=LeftColumn  align=left>&nbsp;" & objCollection.Item.LangLabel.LangLabel & "</TD>" & Chr(13)
				strHTML = strHTML & "	<TD class=LeftColumn  align=left>&nbsp;" & objCollection.Item.LangText & "</TD>" & Chr(13)
				strHTML = strHTML & "	<TD class=RightColumn  align=center><A href=""admLanguageString.asp?ID=" & objCollection.Item.ID & """><IMG src=""Images/Pencil.gif"" alt=""Edit"" border=""0""></A></TD>" & Chr(13)
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
				<TD colspan=5 align=center><%=Lang(strTitle)%></TD>
   			</TR>
		    <TR class=ColumnHeading>
				<TD class=LeftColumn align=left width="40%">&nbsp;<%=Lang("Label")%></TD>
				<TD class=LeftColumn align=left width="40%">&nbsp;<%=Lang("String")%></TD>
				<TD class=RightColumn align=center width="20%"><%=Lang("Options")%></TD>
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
			Response.Write DisplayPageNumbers( "admLanguageStringList.asp?Language=" & intLangID & "&", intPage, intPages )
			%>
		    <TR>
				<TD align=right>
					<BR><INPUT style="WIDTH: 80px; BACKGROUND-COLOR: white" OnClick="VBScript:window.navigate 'admLanguageString.asp?Language=<%=intLangID%>'" id=btnNew name=btnNew type=submit value="<%=Lang("New")%>">
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
