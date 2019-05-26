<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: admOrganisation.asp
'  Date:     $Date: 2004/03/10 08:21:12 $
'  Version:  $Revision: 1.5 $
'  Purpose:  Administration page for creating/modifing Organisations
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
	Dim binUserPermMask, binRequiredPerm
	Dim blnSave, blnIsActive
    Dim objOrg, objCollection
	Dim strOrgName, strOrgShortName, strPhone, strFax, strEMail, strHeading
	Dim strOfficeLocation, strMailAddress, strCourierAddress, strCity, strState
	Dim strIsActiveHTML, strMode, strNotes, strPassword, strCountry
	Dim intPrimaryContactID, intOrgID, intUserID, intLastUpdateByID, intOrgTypeID
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

		intOrgID = CInt(Request.Form("tbxOrgID"))
		strOrgName = Request.Form("tbxOrgName")
		strOrgShortName = Request.Form("tbxOrgShortName")
		strPhone  = Request.Form("tbxPhone")
		strFax  = Request.Form("tbxFax")
		strEMail = Request.Form("tbxEMail")
		strOfficeLocation = Request.Form("txtOfficeLocation")
		strMailAddress = Request.Form("txtMailAddress")
		strCourierAddress = Request.Form("txtCourierAddress")
		strCity = Request.Form("tbxCity")
		strState = Request.Form("tbxState")
		strNotes = Request.Form("txtNotes")
		strPassword = Request.Form("tbxPassword")
		intPrimaryContactID = CInt(Request.Form("cbxPrimaryContact"))
		strCountry = Request.Form("tbxCountry")
		intOrgTypeID = CInt(Request.Form("cbxOrgType"))

		If Request.Form("chkIsActive") = "on" Then
			blnIsActive = lhd_True
		Else
			blnIsActive = lhd_False
		End If

		dteLastUpdate = SQLDateFormat(Day(Now()), Month(Now()), Year(Now())) & " " & FormatDateTime(Now(), vbShortTime)
		intLastUpdateByID = intUserID

		
		' Check for required fields
		
		If Len(strOrgName) = 0 Then
		
			Call DisplayError(1, "All required fields need to be entered, please go back and populate these fields")
			
		Else
		
			' Do nothing
			
		End If


		' Now save/update the Organisation details

		Set objOrg = New clsOrganisation

		objOrg.ID = intOrgID

		If Not objOrg.Load Then
		
			' Organisation does not exist
			Response.Write objOrg.LastError & "<P>"

		Else
		
			' Organisation exists
		
		End If

		' Check the the fields and leave Null if nothing is set.

		objOrg.OrgName = strOrgName
		objOrg.OrgShortName = strOrgShortName
		objOrg.Phone = strPhone
		objOrg.Fax = strFax
		objOrg.EMail = strEMail
		objOrg.OfficeLocation = strOfficeLocation
		objOrg.MailAddress = strMailAddress
		objOrg.CourierAddress = strCourierAddress
		objOrg.City = strCity
		objOrg.State = strState
		objOrg.Notes = strNotes
		objOrg.Password = strPassword
		objOrg.PrimaryContactID = intPrimaryContactID
		objOrg.Country = strCountry
		objOrg.OrgTypeID = intOrgTypeID
		objOrg.IsActive = blnIsActive
		objOrg.LastUpdate = CDate(dteLastUpdate)
		objOrg.LastUpdateByID = intLastUpdateByID
						
		If Not objOrg.Update Then
						
			' Failed to create/save Organisation
			Response.Write objOrg.LastError & "<P>"
							
		Else
						
			intOrgID = objOrg.ID
			strHeading = "Organisation Saved"
						
		End If
						
		Set objOrg = Nothing
%>
		<TR>
		   <TD>
		      <TABLE class=Normal cellSpacing=0 cellPadding=1 width="100%" border=0 bgColor=white>
		         <TR class="lhd_Heading1" >
 				 <TD colspan=5 align=middle><%=strHeading%></TD>
				 </TR>
		      <TR>
		        <TD align=right colspan=5><a style="FONT-SIZE: 8pt" href="admOrganisationList.asp?Page=1"><%=Lang("Manage_Organisations")%></a></TD>
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
 				 <TD colspan=3 align=left>Organisation information has been successfully saved.</TD>
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
	
		' Mode: 1 - To create a new Organisation
		'		2 - Edit a Organisation
	
'		strMode = Request.QueryString("mode")
	
		If Request.QueryString.Count = 0 Then
	
			' Create a new record
	
			strMode = 1
			intOrgID = 0
			strHeading = Lang("New") & " " & Lang("Organisation")
			
		Else
	
			' Edit a record determine by the DeptID passed via the QueryString
	
			strMode = 2
			intOrgID = Request.QueryString("id")
			strHeading = Lang("Modify") & " " & Lang("Organisation")
			
		End If
			

		Select Case strMode
		
			Case 1	' Create new Organisation
			
				strOrgName =""
				strOrgShortName = ""
				strPhone  = ""
				strFax  = ""
				strEMail = ""
				strOfficeLocation = ""
				strMailAddress = ""
				strCourierAddress = ""
				strCity = ""
				strState = ""
				strNotes = ""
				strPassword = ""
				intPrimaryContactID = 0
				strCountry = "Australia"
				intOrgTypeID = 0

				blnIsActive = lhd_True
				strIsActiveHTML = "CHECKED"

				dteLastUpdate = ""
				intLastUpdateByID = 0


			Case 2  ' Edit Organisation

				' Get the Organisation ID we want to edit and load the record

				Set objOrg = New clsOrganisation

				objOrg.ID = intOrgID
			
				If Not objOrg.Load Then
				
					' Couldn't load user for some reason
					Response.Write objOrg.LastError & "<P>"
				
				Else
				
					strOrgName = objOrg.OrgName
					strOrgShortName = objOrg.OrgShortName
					strPhone = objOrg.Phone
					strFax = objOrg.Fax
					strEMail = objOrg.EMail
					strOfficeLocation = objOrg.OfficeLocation
					strMailAddress = objOrg.MailAddress
					strCourierAddress = objOrg.CourierAddress
					strCity = objOrg.City
					strState = objOrg.State
					strNotes = objOrg.Notes
					strPassword = objOrg.Password
					intPrimaryContactID = objOrg.PrimaryContactID
					intOrgTypeID = objOrg.OrgTypeID
					strCountry = objOrg.Country
					
					If objOrg.IsActive = True Then
						strIsActiveHTML = "CHECKED"
					Else
						strIsActiveHTML = ""
					End If
					
					dteLastUpdate = objOrg.LastUpdate
					intLastUpdateByID = objOrg.LastUpdateByID

				End If

				Set objOrg = Nothing
				

			Case Else
				' Do nothing
				
		End Select

%>

  <TR>
    <TD>
	  <TABLE class=Normal cellSpacing=0 cellPadding=1 width="100%" border=0 bgColor=white>
	  <FORM action="admOrganisation.asp" method="post" id=frmOrg name=frmOrg>
	  <INPUT id=tbxOrgID name=tbxOrgID type=hidden value="<%=intOrgID%>">
	  <INPUT id=tbxSave name=tbxSave type=hidden value="1">
	  
		<TR class="lhd_Heading1">
			<TD colspan=5 align=center><%=strHeading%></TD>
		</TR>
		<TR>
		  <TD align=right colspan=5><a style="FONT-SIZE: 8pt" href="admOrganisationList.asp?Page=1"><%=Lang("Manage_Organisations")%></a></TD>
		</TR>
		<TR>
		  <TD width="22%"><%=Lang("Organisation_Type")%>:</TD>
		  <TD width="25%">
		     <SELECT id=cbxOrgType style="WIDTH: 100%" name=cbxOrgType>
		        <%
		        Response.Write BuildList( "ORG_TYPE_LIST", intOrgTypeID )
		        %>
		     </SELECT>
		  </TD>
		  <TD width="5%"></TD>
		  <TD width="18%"><%=Lang("Password")%>:</TD>
		  <TD width="25%"><INPUT id=tbxPassword style="WIDTH: 100%" name=tbxPassword value="<%=strPassword%>"></TD>
		</TR>
		<TR>
		  <TD><%=Lang("Organisation_Short_Name")%>:</TD>
		  <TD><INPUT id=tbxOrgShortName style="WIDTH: 100%" name=tbxOrgShortName value="<%=strOrgShortName%>"></TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		</TR>
		<TR>
		  <TD><%=Lang("Organisation_Name")%>:</TD>
		  <TD><INPUT id=tbxOrgName style="WIDTH: 100%" name=tbxOrgName value="<%=strOrgName%>"></TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		</TR>
		<TR>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		</TR>
		<TR>
		  <TD><%=Lang("Primary_Contact")%>:</TD>
		  <TD>
		     <SELECT id=cbxPrimaryContact style="WIDTH: 100%" name=cbxPrimaryContact>
				<%
				Set objCollection = New clsCollection
					
				objCollection.CollectionType = objCollection.clContact
				objCollection.Query = "SELECT ContactPK, UserName, OrgFK FROM tblContacts WHERE IsActive=" & lhd_True & " And OrgFK =" & intOrgID & " ORDER BY UserName ASC"

				If Not objCollection.Load Then
						
				    Response.Write objCollection.LastError
							    
				Else
								
					If objCollection.BOF And objCollection.EOF Then
								
						' No records
								
					Else
						
						Do While Not objCollection.EOF

				 			If objCollection.Item.ID = intPrimaryContactID Then
							%>
								<OPTION SELECTED VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.UserName%></OPTION>
							<%
							Else
							%>
								<OPTION VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.UserName%></OPTION>
							<%
							End If
							
							objCollection.MoveNext
							
						Loop
									
					End If
								
				End If
					
				Set objCollection = Nothing
				%>
		     </SELECT>
		  </TD>
		  <TD></TD>
		  <TD><%=Lang("Email")%>:</TD>
		  <TD><INPUT id=tbxEMail style="WIDTH: 100%" name=tbxEMail value="<%=strEMail%>"></TD>
		</TR>
		<TR>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		</TR>
		<TR>
		  <TD style="VERTICAL-ALIGN: top"><%=Lang("Office_Address")%>:</TD>
		  <TD><TEXTAREA id=txtOfficeLocation style="WIDTH: 100%" name=txtOfficeLocation><%=strOfficeLocation%></TEXTAREA></TD>
		  <TD></TD>
		  <TD> </TD>
		  <TD></TD>
		</TR>
		<TR>
		  <TD><%=Lang("City")%>:</TD>
		  <TD><INPUT id=tbxCity style="WIDTH: 100%" name=tbxCity value="<%=strCity%>"></TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		</TR>
		<TR>
		  <TD><%=Lang("State")%>:</TD>
		  <TD><INPUT id=tbxState style="WIDTH: 100%" name=tbxState value="<%=strState%>"></TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		</TR>
		<TR>
		  <TD><%=Lang("Country")%>:</TD>
		  <TD><INPUT id=tbxCountry style="WIDTH: 100%" name=tbxCountry value="<%=strCountry%>"></TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		</TR>
		<TR>
		  <TD>  &nbsp; </TD>
		  <TD></TD>
		  <TD></TD>
		  <TD> </TD>
		  <TD></TD>
		</TR>
		<TR>
		  <TD><%=Lang("Phone")%>:</TD>
		  <TD><INPUT id=tbxPhone style="WIDTH: 100%" name=tbxPhone value="<%=strPhone%>"></TD>
		  <TD></TD>
		  <TD><%=Lang("Fax")%>:</TD>
		  <TD><INPUT id=tbxFax style="WIDTH: 100%" name=tbxFax value="<%=strFax%>"></TD>
		</TR>
		<TR>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		</TR>
		<TR>
		  <TD style="VERTICAL-ALIGN: top"><%=Lang("Mail_Address")%>:</TD>
		  <TD><TEXTAREA id=txtMailAddress style="WIDTH: 100%; HEIGHT: 70px" name=txtMailAddress><%=strMailAddress%></TEXTAREA></TD>
		  <TD></TD>
		  <TD style="VERTICAL-ALIGN: top"><%=Lang("Courier_Address")%>:</TD>
		  <TD><TEXTAREA id=txtCourierAddress style="WIDTH: 100%; HEIGHT: 70px" name=txtCourierAddress><%=strCourierAddress%></TEXTAREA></TD>
		</TR>
		<TR>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		</TR>
		<TR>
		  <TD><%=Lang("Notes")%>:</TD>
		  <TD colspan=4><TEXTAREA id=txtNotes style="WIDTH: 100%" name=txtNotes><%=strNotes%></TEXTAREA></TD>
		</TR>
		<TR>
		  <TD></TD>
		  <TD><INPUT id=chkIsActive name=chkIsActive type=checkbox <%=strIsActiveHTML%>>&nbsp;<%=Lang("Is_Active")%></TD>
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
  </TABLE></P></BODY></HTML>
<%
cnnDB.Close
Set cnnDB = Nothing
%>