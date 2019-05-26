<%
' ----------------------------------------------------------------------------------
'
'  Huron Support Desk, Copyright (C) 2003
'
'  File Name:	inoutContact.asp
'
'  Author(s):	t_klose@hotmail.com, etc.
'  Created:		13th Jun 2003
'  Version:		$
'
'  Purpose:		
'
' ----------------------------------------------------------------------------------
%>
<% Option Explicit
%>
<!--METADATA TYPE="TypeLib" NAME="Microsoft ActiveX Data Objects 2.5 Library" UUID="{00000205-0000-0010-8000-00AA006D2EA4}" VERSION="2.5"-->
<!--METADATA TYPE="TypeLib" NAME="Microsoft Scripting Runtime" UUID="{420B2830-E718-11CF-893D-00A0C9054228}" VERSION="1.0"-->
<!--METADATA TYPE="TypeLib" NAME="Microsoft CDO for Windows 2000 Library" UUID="{CD000000-8B95-11D1-82DB-00C04FB1625D}" VERSION="1.0"-->

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


<HTML>
<HEAD>

<META content="MSHTML 6.00.2600.0" name=GENERATOR></HEAD>

<LINK rel="stylesheet" type="text/css" href="Default.css">
<%
  Dim cnnDB
  Dim blnSave, blnIsActive
  Dim objContact, objCollection
  Dim dtIOStatusDate, dtCreated, dteLastUpdate, dtLastAccess
  Dim binPermMask, binUserPermMask, binRequiredPerm

  Dim strFName, strLName, strOfficeLocation, strPagerEmail, strDepartment
  Dim strJobTitle, strOfficePhone, strHomePhone, strMobilePhone, strCompany
  Dim strMode, strUserName, strIOStatusText, strResume, strEmail, strNotes
  Dim strIsActiveHTML, strHeading, strPassword, strDeptName

  Dim intUserID, intPhotoFileID, intTZOffset, intLastUpdateByID, intJobFunctionID
  Dim intContactTypeID, intRoleID, intDeptID, intOrgID, intLangID, intContactID, intIOStatusID
  Dim intPinNumber


  ' Get user variables

  Set cnnDB = CreateConnection
    
  intUserID = GetUserID
  binUserPermMask = GetUserPermMask

%>	
<BODY>
<P align=center>
<TABLE class=Normal align=center cellSpacing=0 cellPadding=0 width="680" border=0>
<TR><TD><% Response.Write DisplayHeader %></TD></TR>
<%

	intContactID = Request.QueryString("id")
	Set objContact = New clsContact
	objContact.ID = intContactID
	
	If Not objContact.Load Then
		' Couldn't load user for some reason
	Else
		strUserName = objContact.UserName
		strPassword = objContact.Password
		strFName = objContact.FName
		strLName = objContact.LName
		intContactTypeID = objContact.ContactTypeID
		intOrgID = objContact.OrgID
		intDeptID = objContact.DeptID
		strDeptName = objContact.Dept.DeptName
		intLangID = objContact.LangID
		intRoleID = objContact.RoleID
		strHomePhone = objContact.HomePhone
		strOfficePhone = objContact.OfficePhone
		strMobilePhone = objContact.MobilePhone
		strJobTitle = objContact.JobTitle
		If IsNull(objContact.JobFunction) Then
			intJobFunctionID = 0
		Else
			intJobFunctionID = objContact.JobFunction
		End If
		strEmail = objContact.Email
		strPagerEmail = objContact.PagerEmail
		strOfficeLocation = objContact.OfficeLocation
		strNotes = objContact.Notes
		intIOStatusID = objContact.IOStatusID
		dtIOStatusDate = objContact.IOStatusDate
		strIOStatusText = objContact.IOStatusText
		intTZOffset = objContact.TZOffset
		binPermMask = objContact.FName
	        if IsNull(objContact.IOStatusID) then
			intPinNumber = 1
	        else
	        	intPinNumber = objContact.IOStatusID - 16 '** Change
	        end if
	
		If objContact.IsActive = True Then
			strIsActiveHTML = "CHECKED"
		Else
			strIsActiveHTML = ""
		End If
	
		dtLastAccess = objContact.LastAccess
		strResume  = objContact.sResume
		intPhotoFileID = objContact.PhotoFileID
		dteLastUpdate = objContact.LastUpdate
		intLastUpdateByID = objContact.LastUpdateByID
	End If
	Set objContact = Nothing
	strHeading = Lang("Modify") & " " & Lang("Contact")
%>
<TR><TD>
	<table class="normal" cellSpacing=0 cellPadding=0 width="100%">
		<tr class="lhd_Heading1"><td colspan=6 align=center><%=lang("UserDetails")%></td></tr>
		<tr><td width=150><%=lang("Name")%></td><td><b><%=strFName & " " & strLName%></b></td><td>&nbsp;</td></tr>
		<tr><td width=150><%=lang("Department")%></td><td><b><%=strDeptName%></b></td><td rowspan=6 valign=top align=right><img src="images/nopicture.gif" alt="User Picture"></td></tr>
		<tr><td width=150><%=lang("User_Name")%></td><td><b><%=strUserName%></b></td></tr>
		<tr><td width=150><%=lang("Email")%></td><td><b><%=strEmail%></b></td></tr>
		<tr><td width=150><%=lang("Phone_Work")%></td><td><b><%=strOfficePhone%></b></td></tr>
		<tr><td width=150><%=lang("Phone_Mobile")%></td><td><b><%=strMobilePhone%></b></td></tr>
		<tr><td width=150><%=lang("Phone_Home")%></td><td><b><%=strHomePhone%></b></td></tr>
		<tr><td width=150><%=lang("Job_Title")%></td><td><%=strJobTitle%></td><td>&nbsp;</td></tr>
		<tr><td width=150><%=lang("Resume")%></td><td><%=strResume%></td><td>&nbsp;</td></tr>
		<tr><td width=150><%=Lang("In/Out_Status")%>:</td><td><img src="images/pin<%=intPinNumber%>.gif" border="" alt="">&nbsp;<%=strIOStatusText%></td><td>&nbsp;</td></tr>
		<tr><td width=150><%=Lang("In/Out_Status_Date")%>:</td><td><%=dtIOStatusDate%></td><td>&nbsp;</td></tr>
		<tr><td colspan=3 align=right><a href="inoutStatus.asp?id=<%=intContactID%>"><%=lang("Change Status")%></a></td></tr>
	</table>
</TD></TR>
<TR><TD><% Response.Write DisplayFooter %></TD></TR>
</TABLE></P>
</BODY>
</HTML>
	
<%
cnnDB.Close
Set cnnDB = Nothing
%>