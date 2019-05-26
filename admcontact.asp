<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: admContact.asp
'  Date:     $Date: 2004/03/11 06:08:42 $
'  Version:  $Revision: 1.8 $
'  Purpose:  Administration page for creating/modifing Contacts
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
	Dim strIsActiveHTML, strHeading, strPassword

  Dim intUserID, intPhotoFileID, intTZOffset, intLastUpdateByID, intJobFunctionID
  Dim intContactTypeID, intRoleID, intDeptID, intOrgID, intLangID, intContactID, intIOStatusID


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

<%
	
	If Request.Form("tbxSave") = "1" Then
		blnSave = True
	Else
		blnSave = False
	End If
	

	If blnSave = True Then	' Start Save
		intContactID = Request.Form("tbxContactID")
		strUserName = Request.Form("tbxUserName")
		strPassword = Request.Form("pwdPassword")
		strFName = Request.Form("tbxFName")
		strLName = Request.Form("tbxLName")
		intContactTypeID = CInt(Request.Form("cbxContactType"))
		intOrgID = CInt(Request.Form("cbxOrganisation"))
		intDeptID = Cint(Request.Form("cbxDepartment"))
		intLangID = CInt(Request.Form("cbxLanguage"))
		intRoleID = CInt(Request.Form("cbxRole"))
		strHomePhone = Request.Form("tbxHomePhone")
		strOfficePhone = Request.Form("tbxOfficePhone")
		strMobilePhone = Request.Form("tbxMobilePhone")
		strJobTitle = Request.Form("tbxJobTitle")
		intJobFunctionID = CInt(Request.Form("cbxJobFunction"))
		strEmail = Request.Form("tbxEmail")
		strPagerEmail = Request.Form("tbxPagerEmail")
		strOfficeLocation = Request.Form("tbxOfficeLocation")
		strNotes = Request.Form("txtNotes")

		If Request.Form("chkIsActive") = "on" Then
			blnIsActive = lhd_True
		Else
			blnIsActive = lhd_False
		End If

		intIOStatusID = CInt(Request.Form("cbxIOStatus"))
		dtIOStatusDate = SQLDate( Request.Form("tbxIOStatusDate") )
		strIOStatusText = Request.Form("tbxIOStatusText")
		intTZOffset = CInt(Request.Form("cbxTZOffset"))
		
'		binPermMask = Request.Form("tbxUserPermMask")
		binPermMask = &H0000000000000000
		
		strResume  = Request.Form("txtResume")
		intPhotoFileID = CInt(Request.Form("cbxPhotoFile"))
		dteLastUpdate = SQLDateFormat(Day(Now()), Month(Now()), Year(Now())) & " " & FormatDateTime(Now(), vbShortTime)
		intLastUpdateByID = intUserID
		
		' Need to check for required fields
		
		If Len(strUserName) = 0 Or Len(strFName) = 0 Or Len(strLName) = 0 Or Len(strEmail) = 0 Or intDeptID = 0 Then
		
			Call DisplayError(1, "All required fields need to be entered, please go back and populate these fields")
			
		End If


		' Now save/update the contacts details

		Set objContact = New clsContact

		objContact.ID = intContactID

		If Not objContact.Load Then
		
			' Contact does not exist

		Else
		
			' Contact exists
		
		End If
		
		objContact.UserName = strUserName
		objContact.FName = strFName
		objContact.LName = strLName

		If intContactTypeID > 0 Then
			objContact.ContactTypeID  = intContactTypeID
		End If
		If intOrgID > 0 Then
			objContact.OrgID = intOrgID
		End If

		objContact.DeptID = intDeptID
		objContact.LangID = intLangID
		objContact.RoleID = intRoleID

		If Len(strHomePhone) > 0 Then
			objContact.HomePhone = strHomePhone
		End If
		If Len(strOfficePhone) > 0 Then
			objContact.OfficePhone = strOfficePhone
		End If
		If Len(strMobilePhone) > 0 Then
			objContact.MobilePhone = strMobilePhone
		End If
		If Len(strJobTitle) > 0 Then
			objContact.JobTitle = strJobTitle
		End If
		If intJobFunctionID > 0 Then
			objContact.JobFunctionID = intJobFunctionID
		End If

		objContact.Email = strEmail
		objContact.Password = strPassword

		If Len(strPagerEmail) > 0 Then
			objContact.PagerEmail = strPagerEmail
		End If
		If Len(strOfficeLocation) > 0 Then
			objContact.OfficeLocation = strOfficeLocation
		End If
		If Len(strNotes) > 0 Then
			objContact.Notes = strNotes
		End If

		objContact.IsActive = blnIsActive

		If intIOStatusID > 0 Then
			objContact.IOStatusID = intIOStatusID
		End If
		If IsDate(dtIOStatusDate) Then
			objContact.IOStatusDate = dtIOStatusDate
		End If
		If Len(strIOStatusText) > 0 Then
			objContact.IOStatusText = strIOStatusText
		End If
		If intTZOffset > 0 Then
			objContact.TZOffset = intTZOffset
		End If
'		objContact.UserPermMask = binPermMask
		If Len(strResume) > 0 Then
			objContact.sResume = strResume
		End If
		If intPhotoFileID > 0 Then
			objContact.PhotoFileID = intPhotoFileID
		End If
		If IsDate(dteLastUpdate) Then
			objContact.LastUpdate = dteLastUpdate
		End If
		If intLastUpdateByID > 0 Then
			objContact.LastUpdateByID = intLastUpdateByID
		End If
						
		If Not objContact.Update Then
						
			' Failed to create/save user
							
		Else
						
			intContactID = objContact.ID
			strHeading = Lang("Contact_Saved")
						
		End If
						
		Set objContact = Nothing
		
		
%>
		<TR>
		   <TD>
		      <TABLE class=Normal width="100%" border=0 cellSpacing=0 cellPadding=1>
				<TR class="lhd_Heading1">
				   <TD colspan=5 align=center><%=strHeading%></TD>
				</TR>
		    <TR>
		      <TD align=right colspan=5><a style="FONT-SIZE: 8pt" href="admContactList.asp?Page=1"><%=Lang("Manage_Contacts")%></a></TD>
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
 				   <TD colspan=3 align=Left>Contact information has been successfully saved.</TD>
 				   <TD></TD>
				 </TR>
				<TR>
				   <TD colspan=5></TD>
				</TR>
		      </TABLE>
		   </TD>
		</TR>
<%
	Else
	
		If Request.QueryString.Count = 0 Then
	
			' Create a new record
	
			intContactID = 0
			strMode = 1
			strHeading = Lang("New_Contact")
			
		Else
	
			' Edit a record determine by the Contact ID passed via the QueryString
	
			intContactID = CInt(Request.QueryString("ID"))
			strMode = 2
			strHeading = Lang("Modify_Contact")
			
		End If


		
		Select Case strMode
		
			Case 1	' New Contact
			
				strUserName = ""
				strFName = ""
				strLName = ""
				intContactTypeID = Application("DEFAULT_CONTACT_TYPE")
				strCompany = ""
				intOrgID = 0
				intDeptID = 0
				strDepartment = ""
				intLangID = Application("DEFAULT_LANGUAGE")
				intRoleID = Application("DEFAULT_ROLE")
				strHomePhone = ""
				strOfficePhone = ""
				strMobilePhone = ""
				strJobTitle = ""
				intJobFunctionID = 0
				strEmail = ""
				strPagerEmail = ""
				strOfficeLocation = ""
				strNotes = ""
				intIOStatusID = 0
				dtIOStatusDate = ""
				strIOStatusText = ""
				intTZOffset = 0
				binPermMask = ""
				blnIsActive = lhd_True
				strIsActiveHTML = "CHECKED"
				strResume  = ""
				intPhotoFileID = 0
				dtLastAccess = ""
				dteLastUpdate = ""
				intLastUpdateByID = 0

				strHeading = Lang("New_Contact")

			Case 2  ' Edit Contact

				Set objContact = New clsContact

				objContact.ID = intContactID
				'UserName = Right(Request.ServerVariables("AUTH_USER"), Len(Request.ServerVariables("AUTH_USER")) - InStr(Request.ServerVariables("AUTH_USER"), "\"))
			
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

			Case Else
				' Do nothing
				
		End Select
	

%>

  <TR>
    <TD>
	  <TABLE class=Normal cellSpacing=0 cellPadding=1 width="100%" border=0 bgColor=white>
	  <FORM action="admContact.asp" method="post" id=frmContact name=frmContact>
	  <INPUT id=tbxSave name=tbxSave type=hidden value="1">
	  <INPUT id=tbxContactID name=tbxContactID type=hidden value="<%=intContactID%>">
	  

	  
		<TR class="lhd_Heading1">
			<TD colspan=5 align=center><%=strHeading%></TD>
		</TR>
		<TR>
		  <TD align=right colspan=5><a style="FONT-SIZE: 8pt" href="admContactList.asp?Page=1"><%=Lang("Manage_Contacts")%></a></TD>
		</TR>
  <TR>
    <TD width="22%"><B><%=Lang("User_Name")%>:</B></TD>
    <TD width="25%"><INPUT id=tbxUserName name=tbxUserName style="WIDTH: 100%" value="<%=strUserName%>"></TD>
    <TD width="5%"></TD>
    <TD width="18%"><%=Lang("Last_Access_Time")%>:</TD>
    <TD width="25%"><%=DisplayDateTime(dtLastAccess)%></TD>
  </TR>
  <TR>
    <TD><%=Lang("Password")%>:</TD>
    <TD><INPUT id=pwdPassword name=pwdPassword type=password style="WIDTH: 100%" value="<%=strPassword%>"></TD>
    <TD></TD>
    <TD></TD>
    <TD></TD></TR>
  <TR height=27>
    <TD><B><%=Lang("First_Name")%>:</TD>
    <TD><INPUT id=tbxFName name=tbxFName style="WIDTH: 100%" value="<%=strFName%>"
         ></TD>
    <TD style="WIDTH: 50px" width=50></TD>
    <TD><B><%=Lang("Last_Name")%>:</B></TD>
    <TD><INPUT id=tbxLName name=tbxLName style="WIDTH: 100%"  value="<%=strLName%>"
         ></TD></TR>
  <TR height=27>
    <TD><B><%=Lang("Email")%>:</B></TD>
    <TD><INPUT id=tbxEmail name=tbxEmail 
            style="WIDTH: 100%" value="<%=strEmail%>"></TD>
    <TD></TD>
    <TD></TD>
    <TD></TD></TR>
  <TR>
    <TD>
            <P align=left><%=Lang("Pager_Email")%>:</P></TD>
    <TD><INPUT id=tbxPagerEmail name=tbxPagerEmail 
            style="WIDTH: 100%" value="<%=strPagerEmail%>"></TD>
    <TD></TD>
    <TD></TD>
    <TD></TD></TR>
  <TR>
    <TD><%=Lang("Contact_Type")%>:</TD>
    <TD>
    <SELECT id=cbxContactType name=cbxContactType style="WIDTH: 100%">
	   <OPTION SELECTED VALUE="0">(None)</OPTION>
		<%
		Response.Write BuildList("CONTACT_TYPE_LIST", intContactTypeID)
		%>
     </SELECT>
     </TD>
    <TD></TD>
    <TD></TD>
    <TD></TD></TR>
  <TR>
    <TD><%=Lang("Organisation")%>:</TD>
    <TD>
       <SELECT id=cbxOrganisation name=cbxOrganisation style="WIDTH: 100%">
	   <OPTION SELECTED VALUE="0">(None)</OPTION>
	   <%
		Set objCollection = New clsCollection
	   
		objCollection.CollectionType = objCollection.clOrganisation
		objCollection.Query = "SELECT OrgPK, OrgName FROM tblOrganisations WHERE IsActive=" & lhd_True & " ORDER BY OrgPK ASC"
								
		If Not objCollection.Load Then
		
			Response.Write objCollection.LastError
			
		Else
		
		    Do While Not objCollection.EOF
		    
				If objCollection.Item.ID = intOrgID Then
				%>
					<OPTION SELECTED VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.OrgName%></OPTION>
				<%
				Else
				%>
					<OPTION VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.OrgName%></OPTION>
				<%
				End If
				
				objCollection.MoveNext
				
		    Loop
		    
		End If
							
		Set objCollection = Nothing
		%>
       </SELECT>
    </TD>
    <TD></TD>
    <TD></TD>
    <TD></TD></TR>
  <TR>
    <TD class="lhd_Required"><%=Lang("Department")%>:</TD>
    <TD>
    <SELECT id=cbxDepartment name=cbxDepartment style="WIDTH: 100%">
	   <OPTION SELECTED VALUE="0">(None)</OPTION>
	   <%
		Set objCollection = New clsCollection
	   
		objCollection.CollectionType = objCollection.clDepartment
		objCollection.Query = "SELECT DeptPK, DeptName FROM tblDepartments WHERE IsActive=" & lhd_True & " ORDER BY DeptPK ASC"
								
		If Not objCollection.Load Then
		
			Response.Write objCollection.LastError
			
		Else
		
		    Do While Not objCollection.EOF
		    
				If objCollection.Item.ID = intDeptID Then
				%>
					<OPTION SELECTED VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.DeptName%></OPTION>
				<%
				Else
				%>
					<OPTION VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.DeptName%></OPTION>
				<%
				End If
				
				objCollection.MoveNext
				
		    Loop
		    
		End If
							
		Set objCollection = Nothing
		%>
         </SELECT>
         </TD>
    <TD></TD>
    <TD><%=Lang("Location")%>:</TD>
    <TD><INPUT id=tbxOfficeLocation name=tbxOfficeLocation 
            style="WIDTH: 100%" value="<%=strOfficeLocation%>"></TD></TR>
  <TR>
    <TD><%=Lang("Job_Title")%>:</TD>
    <TD><INPUT id=tbxJobTitle name=tbxJobTitle 
            style="WIDTH: 100%" value="<%=strJobTitle%>"></TD>
    <TD></TD>
    <TD><%=Lang("Job_Function")%>:</TD>
    <TD>
    <SELECT id=cbxJobFunction name=cbxJobFunction style="WIDTH: 100%">
	   <OPTION SELECTED VALUE="0">(None)</OPTION>
		<%
		Response.Write BuildList("JOB_FUNCTION_LIST", intJobFunctionID)
		%>
	</SELECT></TD></TR>
  <TR>
    <TD class="lhd_Required"><%=Lang("Phone_Work")%>:</TD>
    <TD><INPUT id=tbxOfficePhone name=tbxOfficePhone 
            style="WIDTH: 100%" value="<%=strOfficePhone%>"></TD>
    <TD></TD>
    <TD></TD>
    <TD></TD></TR>
  <TR>
    <TD><%=Lang("Phone_Home")%>:</TD>
    <TD><INPUT id=tbxHomePhone name=tbxHomePhone 
            style="WIDTH: 100%" value="<%=strHomePhone%>"></TD>
    <TD></TD>
    <TD> </TD>
    <TD></TD></TR>
  <TR>
    <TD><%=Lang("Phone_Mobile")%>:</TD>
    <TD><INPUT id=tbxMobilePhone name=tbxMobilePhone style="WIDTH: 100%" value="<%=strMobilePhone%>"></TD>
    <TD></TD>
    <TD></TD>
    <TD></TD></TR>
  <TR>
    <TD><B><%=Lang("Language")%>:</B></TD>
    <TD>
	   <SELECT id=cbxLanguage name=cbxLanguage style="WIDTH: 100%">
	   <%
		Set objCollection = New clsCollection
	   
		objCollection.CollectionType = objCollection.clLanguage
		objCollection.Query = "SELECT LangPK, LangName FROM tblLanguages WHERE IsActive=" & lhd_True & " ORDER BY LangPK ASC"
								
		If Not objCollection.Load Then
		
			Response.Write objCollection.LastError
			
		Else
		
		    Do While Not objCollection.EOF
		    
				If objCollection.Item.ID = intLangID Then
				%>
					<OPTION SELECTED VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.LangName%></OPTION>
				<%
				Else
				%>
					<OPTION VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.LangName%></OPTION>
				<%
				End If
				
				objCollection.MoveNext
				
		    Loop
		    
		End If
							
		Set objCollection = Nothing
		%>
       </SELECT>
    </TD>
    <TD></TD>
    <TD></TD>
    <TD></TD></TR>
  <TR>
    <TD vAlign=top><%=Lang("Resume")%>:</TD>
    <TD colspan=4><TEXTAREA id=txtResume name=txtResume style="HEIGHT: 70px; WIDTH: 100%"><%=strResume%></TEXTAREA></TD></TR>
  <TR>
    <TD></TD>
    <TD></TD>
    <TD></TD>
    <TD></TD>
    <TD></TD></TR>
  <TR>
    <TD><B><%=Lang("Role")%>:</B></TD>
    <TD>
    <SELECT id=cbxRole name=cbxRole style="WIDTH: 100%"> 
	   <%
		Set objCollection = New clsCollection
	   
		objCollection.CollectionType = objCollection.clRole
		objCollection.Query = "SELECT RolePK, RoleName FROM tblRoles WHERE IsActive=" & lhd_True & " ORDER BY RolePK ASC"
								
		If Not objCollection.Load Then
		
			Response.Write objCollection.LastError
			
		Else
		
		    Do While Not objCollection.EOF
		    
				If objCollection.Item.ID = intRoleID Then
				%>
					<OPTION SELECTED VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.RoleName%></OPTION>
				<%
				Else
				%>
					<OPTION VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.RoleName%></OPTION>
				<%
				End If
				
				objCollection.MoveNext
				
		    Loop
		    
		End If
							
		Set objCollection = Nothing
		%>
       </SELECT>
       </TD>
    <TD></TD>
    <TD><%=Lang("User_Permissions")%>:</TD>
    <TD><INPUT id=tbxUserPermMask name=tbxUserPermMask style="WIDTH: 100%" value"<%=binPermMask%>"></TD>
  </TR>
  <TR>
    <TD></TD>
    <TD></TD>
    <TD></TD>
    <TD></TD>
    <TD></TD></TR>
  <TR>
    <TD><%=Lang("In/Out_Status")%>:</TD>
    <TD>
    <SELECT id=cbxIOStatus name=cbxIOStatus style="WIDTH: 100%">
   	   <OPTION SELECTED VALUE="0">(None)
   	</OPTION>
	</SELECT></TD>
    <TD></TD>
    <TD><%=Lang("In/Out_Status_Date")%>:</TD>
    <TD><INPUT id=tbxIOStatusDate name=tbxIOStatusDate style="WIDTH: 100%" value"<%=dtIOStatusDate%>"></TD></TR>
  <TR>
    <TD><%=Lang("In/Out_Status_Text")%>:</TD>
    <TD colspan=4><INPUT id=tbxIOStatusText name=tbxIOStatusText style="WIDTH: 100%" value"<%=strIOStatusText%>"></TD></TR>
  <TR>
    <TD></TD>
    <TD></TD>
    <TD></TD>
    <TD></TD>
    <TD></TD></TR>
  <TR>
    <TD><%=Lang("Timezone_Offset")%>:</TD>
    <TD><SELECT id=cbxTZOffset name=cbxTZOffset style="WIDTH: 100%">
            <OPTION SELECTED VALUE="0">(None)</OPTION>
		</SELECT>
	</TD>
	<TD></TD>
    <TD></TD>
    <TD></TD>
  </TR>
  <TR>
    <TD></TD>
    <TD></TD>
    <TD></TD>
    <TD></TD>
    <TD></TD></TR>
  <TR>
    <TD><%=Lang("Photo_File")%>:</TD>
    <TD>
       <SELECT id=cbxPhotoFile name=cbxPhotoFile style="WIDTH: 100%">
	   <OPTION SELECTED VALUE="0">(None)</OPTION>
       </SELECT>
    </TD>
    <TD></TD>
    <TD></TD>
    <TD></TD>
  </TR>
  <TR>
    <TD></TD>
    <TD></TD>
    <TD></TD>
    <TD></TD>
    <TD></TD></TR>
  <TR>
    <TD vAlign=top><%=Lang("Notes")%>:</TD>
    <TD colspan=4><TEXTAREA id=txtNotes name=txtNotes style="HEIGHT: 70px; WIDTH: 100%"><%=strNotes%></TEXTAREA></TD></TR>
  <TR>
    <TD></TD>
    <TD><INPUT id=chkIsActive name=chkIsActive type=checkbox <%=strIsActiveHTML%>>&nbsp;<%=Lang("Is_Active")%></TD>
    <TD></TD>
    <TD></TD>
    <TD></TD></TR>
  <TR>
    <TD></TD>
    <TD></TD>
    <TD></TD>
    <TD></TD>
	<TD align=right><INPUT style="WIDTH: 80px; BACKGROUND-COLOR: white" id=btnSave name=btnSave type=submit value="<%=Lang("Save")%>"></TD>
	</TR>
    </FORM>
  </TABLE>
</TD></TR>
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
  </TABLE></P>
</BODY></HTML>
	
<%
cnnDB.Close
Set cnnDB = Nothing
%>