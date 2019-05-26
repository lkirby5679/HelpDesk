<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: admLanguage.asp
'  Date:     $Date: 2004/03/11 06:08:42 $
'  Version:  $Revision: 1.7 $
'  Purpose:  Administration page for creating/modifing Languages
' ----------------------------------------------------------------------------------
%>
<% 

Option Explicit

%>

<!--METADATA TYPE="TypeLib" NAME="Microsoft ActiveX Data Objects 2.5 Library" UUID="{00000205-0000-0010-8000-00AA006D2EA4}" VERSION="2.5"-->
<!--METADATA TYPE="TypeLib" NAME="Microsoft Scripting Runtime" UUID="{420B2830-E718-11CF-893D-00A0C9054228}" VERSION="1.0"-->

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
Dim objLanguage, objCollection
Dim binUserPermMask, binRequiredPerm
Dim blnSave, blnIsActive, blnIsRTL
Dim strLangName, strISO639, strEncoding, strLocalized
Dim strMode, strIsActive_HTML, strIsRTL_HTML, strHeading
Dim intUserID, intLastUpdateByID, intLanguageID
Dim dteLastUpdate


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

<HTML>
<HEAD>

<META content="MSHTML 6.00.2600.0" name=GENERATOR></HEAD>

<LINK rel="stylesheet" type="text/css" href="Default.css">

<BODY>
<P align=center>
<TABLE class=Normal align=center cellSpacing=1 cellPadding=1 width="680" border=0>
	<TBODY>
	<TR>
		<TD>
		<%
		Response.Write DisplayHeader
		%>
		</TD>
	</TR>

	<%
	
	If Request.Form("tbxSave") = "1" Then
		blnSave = True
	Else
		blnSave = False
	End If
	

	If blnSave = True Then	' Start Save

		intLanguageID = Cint(Request.Form("tbxLanguageID"))
		
		strLangName = Request.Form("tbxLangName")
		strLocalized = Request.Form("tbxLocalized")
		
		If Request.Form("chkIsRTL") = "on" Then
			blnIsActive = lhd_True
		Else
			blnIsActive = lhd_False
		End If

		strEncoding = Request.Form("tbxEncoding")
		strISO639 = Request.Form("tbxISO639")

		If Request.Form("chkIsActive") = "on" Then
			blnIsActive = lhd_True
		Else
			blnIsActive = lhd_False
		End If

		dteLastUpdate = SQLDateFormat(Day(Now()), Month(Now()), Year(Now())) & " " & FormatDateTime(Now(), vbShortTime)
		intLastUpdateByID = intUserID

		
		' Check for required fields
		
		If Len(strLangName) = 0 Then
		
			Call DisplayError(1, "All required fields need to be entered, please go back and populate these fields")
			
		Else
		
			' Do nothing
			
		End If


		' Now save/update the Language details

		Set objLanguage = New clsLanguage

		objLanguage.ID = intLanguageID

		If Not objLanguage.Load Then
			' Language does not exist
		Else
			' Language exists
		End If

		' Check the the fields and leave Null if nothing is set.

		objLanguage.LangName = strLangName
		objLanguage.Localized = strLocalized
		objLanguage.Encoding = strEncoding
		objLanguage.ISO639 = strISO639
		objLanguage.IsRTL = blnIsRTL
		objLanguage.IsActive = blnIsActive
		objLanguage.LastUpdate = dteLastUpdate
		objLanguage.LastUpdateByID = intLastUpdateByID
						
		If Not objLanguage.Update Then
						
			' Failed to create/save Language
			Response.Write objLanguage.LastError
							
		Else
						
			intLanguageID = objLanguage.ID
			strHeading = Lang("Language_Saved")
						
		End If
						
		Set objLanguage = Nothing
		
		%>
		
		<TR>
			<TD>
				<TABLE class=Normal cellSpacing=0 cellPadding=1 width="100%" border=0 bgColor=white>
					<TR class="lhd_Heading1">
						<TD></TD>
 						<TD colspan=3 align=middle><%=strHeading%></TD>
						<TD></TD>
					</TR>
		      <TR>
		        <TD align=right colspan=5><a style="FONT-SIZE: 8pt" href="admLanguageList.asp?Page=1"><%=Lang("Manage_Languages")%></a></TD>
		      </TR>
					<TR>
 						<TD width="10%"></TD>
						<TD width="20%"></TD>
						<TD width="40%"></TD>
						<TD width="20%"></TD>
						<TD width="10%"></TD>
					</TR>
					<TR>
						<TD></TD>
 						<TD colspan=3 align=left>Language information has been successfully saved.</TD>
						<TD></TD>
					</TR>
					<TR>
						<TD></TD>
 						<TD colspan=3></TD>
						<TD></TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
	<%
	Else
	
		' Mode: 1 - To create a new Language
		'		2 - Edit a Language
	
		If Request.QueryString.Count = 0 Then
	
			' Create a new record
	
			strMode = 1
			intLanguageID = 0
			strHeading = Lang("New_Language")
			
		Else
	
			' Edit a record determine by the Language ID passed via the QueryString
	
			strMode = 2
			intLanguageID = Request.QueryString("ID")
			strHeading = Lang("Modify_Language")
			
		End If
			

		Select Case strMode
		
			Case 1	' Create new Language
			
				strLangName = ""
				strLocalized = ""
				strEncoding = ""
				strISO639 = ""

				blnIsRTL = lhd_True
				strIsRTL_HTML = "CHECKED"
				
				blnIsActive = lhd_True
				strIsActive_HTML = "CHECKED"

				dteLastUpdate = ""
				intLastUpdateByID = 0


			Case 2  ' Edit Language

				' Get the Language ID we want to edit and load the record

				Set objLanguage = New clsLanguage

				objLanguage.ID = intLanguageID
			
				If Not objLanguage.Load Then
				
					' Couldn't load Language for some reason
					Response.Write objLanguage.LastError
				
				Else
				
					strLangName = objLanguage.LangName
					strLocalized = objLanguage.Localized
					strEncoding = objLanguage.Encoding
					strISO639 = objLanguage.ISO639

					If objLanguage.IsActive = True Then
						strIsActive_HTML = "CHECKED"
					Else
						strIsActive_HTML = ""
					End If

					If objLanguage.IsRTL = True Then
						strIsRTL_HTML = "CHECKED"
					Else
						strIsRTL_HTML = ""
					End If

					dteLastUpdate = objLanguage.LastUpdate
					intLastUpdateByID = objLanguage.LastUpdateByID

				End If

				Set objLanguage = Nothing
				

			Case Else
				' Do nothing
				
		End Select

	%>

		<TR>
			<TD>
				<TABLE class=Normal cellSpacing=0 cellPadding=1 width="100%" border=0 bgColor=white>
				<FORM action="admLanguage.asp" method="post" id=frmLanguage name=frmLanguage>
				<INPUT id=tbxLanguageID name=tbxLanguageID type=hidden value="<%=intLanguageID%>">
				<INPUT id=tbxSave name=tbxSave type=hidden value="1">
	  
				<TR class="lhd_Heading1">
					<TD colspan=5 align=center><%=strHeading%></TD>
				</TR>
		    <TR>
		      <TD align=right colspan=5><a style="FONT-SIZE: 8pt" href="admLanguageList.asp?Page=1"><%=Lang("Manage_Languages")%></a></TD>
		    </TR>
				<TR>
					<TD width="22%"><B><%=Lang("Language_Name")%>:</B></TD>
					<TD width="25%"><INPUT id=tbxLangName name=tbxLangName style="WIDTH: 100%" value="<%=strLangName%>"></TD>
					<TD width="5%"></TD>
					<TD width="18%"></TD>
					<TD width="25%"></TD>
				</TR>
				<TR>
					<TD><%=Lang("Localised")%>:</TD>
					<TD><INPUT id=tbxLocalized name=tbxLocalized style="WIDTH: 100%" value="<%=strLocalized%>"></TD>
					<TD></TD>
					<TD><INPUT id=chkIsRTL name=chkIsRTL type=checkbox <%=strIsRTL_HTML%>>&nbsp;<%=Lang("Is_RTL")%></TD>
					<TD></TD>
				</TR>
				<TR>
					<TD><%=Lang("Encoding")%>:</TD>
					<TD><INPUT id=tbxEncoding name=tbxEncoding style="WIDTH: 100%" value="<%=strEncoding%>"></TD>
					<TD></TD>
					<TD></TD>
					<TD></TD>
				</TR>
				<TR>
					<TD><%=Lang("ISO_639")%>:</TD>
					<TD><INPUT id=tbxISO639 name=tbxISO639 style="WIDTH: 100%" value="<%=strISO639%>"></TD>
					<TD></TD>
					<TD></TD>
					<TD></TD>
				</TR>
				<TR>
					<TD></TD>
					<TD><INPUT id=chkIsActive name=chkIsActive type=checkbox <%=strIsActive_HTML%>>&nbsp;<%=Lang("Is_Active")%></TD>
					<TD></TD>
					<TD></TD>
					<TD></TD>
				</TR>
				<TR>
					<TD></TD>
					<TD></TD>
					<TD></TD>
					<TD></TD>
					<TD align=right><INPUT style="WIDTH: 80px; BACKGROUND-COLOR: white" id=btnSave name=btnSave type=submit value="<%=Lang("Save")%>"></TD>
				</TR>
			</FORM>
			</TABLE>
		</TD>
	</TR>
	
	<%
	End If	' End Save
	%>
	
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
