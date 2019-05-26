<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: admCaseType.asp
'  Date:     $Date: 2004/03/10 08:21:12 $
'  Version:  $Revision: 1.5 $
'  Purpose:  Administration page for creating/modifing Case Types
' ----------------------------------------------------------------------------------

%>
<%

Option Explicit

%>

<!--METADATA TYPE="TypeLib" NAME="Microsoft ActiveX Data Objects 2.5 Library" UUID="{00000205-0000-0010-8000-00AA006D2EA4}" VERSION="2.5"-->
<!--METADATA TYPE="TypeLib" NAME="Microsoft Scripting Runtime" UUID="{420B2830-E718-11CF-893D-00A0C9054228}" VERSION="1.0"-->
<!--METADATA TYPE="TypeLib" NAME="Microsoft CDO for Windows 2000 Library" UUID="{CD000000-8B95-11D1-82DB-00C04FB1625D}" VERSION="1.0"-->

<!-- #Include File = "Include/Public.asp" -->
<!-- #Include File = "Include/Settings.asp" -->

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
    Dim objCaseType, objCollection
    Dim strCaseTypeName, strCaseTypeDesc, strMode, strIsActiveHTML, strHeading
    Dim intLastUpdateByID, intCaseTypeID, intUserID, intRepGroupID, intCaseTypeOrder
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

		intCaseTypeID = Cint(Request.Form("tbxCaseTypeID"))

		strCaseTypeName = Request.Form("tbxCaseTypeName")
		strCaseTypeDesc = Request.Form("tbxCaseTypeDesc")
		intCaseTypeOrder = CInt(Request.Form("tbxCaseTypeOrder"))
		intRepGroupID = Cint(Request.Form("cbxRepGroup"))

		If Request.Form("chkIsActive") = "on" Then
			blnIsActive = True
		Else
			blnIsActive = lhd_False
		End If

		dteLastUpdate = SQLDateFormat(Day(Now()), Month(Now()), Year(Now())) & " " & FormatDateTime(Now(), vbShortTime)
		intLastUpdateByID = intUserID

		
		' Check for required fields
		
		If Len(strCaseTypeName) = 0 Then
		
			DisplayError 1, "All required fields need to be entered, please go back and populate these fields"
			
		Else
		
			' Do nothing
			
		End If


		' Now save/update the Case Type details

		Set objCaseType = New clsCaseType

		objCaseType.ID = intCaseTypeID

		If Not objCaseType.Load Then
		
			' Case Type does not exist, so we need to create a new one
'			Response.Write objCaseType.LastError

		Else
		
			' Contact exists
		
		End If

		' Check the the fields and leave Null if nothing is set.

		objCaseType.CaseTypeName = strCaseTypeName
		objCaseType.CaseTypeOrder = intCaseTypeOrder
		objCaseType.RepGroupID = intRepGroupID
		
		If Len(strCaseTypeDesc) > 0 Then 
			objCaseType.CaseTypeDesc = strCaseTypeDesc
		End If
		
		objCaseType.IsActive = blnIsActive
		objCaseType.LastUpdate = dteLastUpdate
		objCaseType.LastUpdateByID = intLastUpdateByID

						
		If Not objCaseType.Update Then
						
			' Failed to create/save Case Type
			Response.Write objCaseType.LastError
							
		Else
						
			intCaseTypeID = objCaseType.ID
			strHeading = Lang("Case_Type_Saved")
						
		End If
						
		Set objCaseType = Nothing
%>
		<TR>
		   <TD>
		      <TABLE class=Normal width="100%" border=0 cellSpacing=0 cellPadding=1>
					<TR class="lhd_Heading1">
						<TD colspan=5 align=center><%=strHeading%></TD>
					</TR>
		      <TR>
		        <TD align=right colspan=5><a style="FONT-SIZE: 8pt" href="admCaseTypeList.asp?Page=1"><%=Lang("Manage_Case_Types")%></a></TD>
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
						<TD colspan=3 align="Left">Case Type changes have been successfully saved.</TD>
						<TD></TD>
					</TR>
					<TR>
						<TD colspan="5"></TD>
					</TR>
		      </TABLE>
		   </TD>
		</TR>
<%
	Else
	
		' Mode: 1 - To create a new Case Type
		'		2 - Edit a Case Type
	
'		strMode = Request.QueryString("mode")
	
		If Request.QueryString.Count = 0 Then
	
			' Create a new record
	
			strMode = 1
			intCaseTypeID = 0
			strHeading = Lang("New_Case_Type")
			
		Else
	
			' Edit a record determine by the Case Type ID passed via the QueryString
	
			strMode = 2
			intCaseTypeID = Request.QueryString("id")
			strHeading = Lang("Modify_Case_Type")
			
		End If
			

		Select Case strMode
		
			Case 1	' Create new Case Type
			
				strCaseTypeName = ""
				strCaseTypeDesc = ""
				intCaseTypeOrder = 0

				blnIsActive = lhd_True
				strIsActiveHTML = "CHECKED"

				dteLastUpdate = ""
				intLastUpdateByID = 0


			Case 2  ' Edit Case Type

				' Get the Case Type ID we want to edit and load the record

				Set objCaseType = New clsCaseType

				objCaseType.ID = intCaseTypeID
			
				If Not objCaseType.Load Then
				
					' Couldn't load user for some reason
					Response.Write objCaseType.LastError & "<P>"
				
				Else
				
					strCaseTypeName = objCaseType.CaseTypeName
					strCaseTypeDesc = objCaseType.CaseTypeDesc
					intCaseTypeOrder = CInt(objCaseType.CaseTypeOrder)
					intRepGroupID = CInt(objCaseType.RepGroupID)
					
					If objCaseType.IsActive = True Then
						strIsActiveHTML = "CHECKED"
					Else
						strIsActiveHTML = ""
					End If
					
					dteLastUpdate = objCaseType.LastUpdate
					intLastUpdateByID = objCaseType.LastUpdateByID

				End If

				Set objCaseType = Nothing
				
			Case Else
				' Do nothing
				
		End Select

%>

  <TR>
    <TD>
	  <TABLE class=Normal cellSpacing=0 cellPadding=1 width="100%" border=0 bgColor=white>
	  <FORM action="admCaseType.asp" method="post" id=frmCaseType name=frmCaseType>
	  <INPUT id=tbxCaseTypeID name=tbxCaseTypeID type=hidden value="<%=intCaseTypeID%>">
	  <INPUT id=tbxSave name=tbxSave type=hidden value="1">
		<TR class="lhd_Heading1">
			<TD colspan=5 align=center><%=strHeading%></TD>
		</TR>
		<TR>
		  <TD align=right colspan=5><a style="FONT-SIZE: 8pt" href="admCaseTypeList.asp?Page=1"><%=Lang("Manage_Case_Types")%></a></TD>
		</TR>
		<TR>
		  <TD width="22%"><B><%=Lang("Case_Type")%>:</B></TD>
		  <TD width="25%"><INPUT id=tbxCaseTypeName name=tbxCaseTypeName style="WIDTH: 100%" value="<%=strCaseTypeName%>" ></TD>
		  <TD width="5%"></TD>
		  <TD width="18%"></TD>
		  <TD width="25%"></TD>
		</TR>
		<TR>
		  <TD><%=Lang("Description")%>:</TD>
		  <TD colspan=4><INPUT id=tbxCaseTypeDesc name=tbxCaseTypeDesc 
            style="WIDTH: 100%" value="<%=strCaseTypeDesc%>"></TD>
		</TR>
		<TR>
		  <TD class="lhd_Required"><%=Lang("Rep_Group")%>:</TD>
		  <TD>
			<SELECT id=cbxRepGroup name=cbxRepGroup style="WIDTH: 100%">
				<%
				Set objCollection = New clsCollection

				objCollection.CollectionType = objCollection.clGroup
				objCollection.Query = "SELECT * FROM tblGroups WHERE IsActive=" & lhd_True & " ORDER BY GroupPK ASC"
					    
				If Not objCollection.Load Then
					    
					' Didn't load
							
				Else
				
					Do While Not objCollection.EOF

						If objCollection.Item.ID = intRepGroupID Then
						%>
							<OPTION SELECTED VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.GroupName%></OPTION>
						<%
						Else
						%>
							<OPTION VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.GroupName%></OPTION>
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
		  <TD></TD>
		</TR>
		<TR>
		  <TD><%=Lang("Order")%>:</TD>
		  <TD><INPUT id=tbxCaseTypeOrder name=tbxCaseTypeOrder style="WIDTH: 100%" value="<%=intCaseTypeOrder%>"></TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
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
  </TABLE>
  </P>
  </BODY>
  </HTML>
  
<%
  
	cnnDB.Close
	Set cnnDB = Nothing
  
%>
