<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: admCategory.asp
'  Date:     $Date: 2004/03/17 19:48:08 $
'  Version:  $Revision: 1.6 $
'  Purpose:  Administration page for creating/modifing Categories
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
  Dim objCategory, objCollection
  Dim strCatName, strCatDesc, strMode, strIsActiveHTML, strHeading
  Dim intLastUpdateByID, intCaseTypeID, intUserID, intCatID, intCatOrder
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

		Call ValidateInt(Request.Form("tbxCatID"),-824,"Category ID",intCatID)
		strCatName = Request.Form("tbxCatName")
		strCatDesc = Request.Form("tbxCatDesc")
		Call ValidateInt(Request.Form("tbxCatOrder"),-824,"Category Order",intCatOrder)
		Call ValidateInt(Request.Form("cbxCaseType"),0,"Category Type",intCaseTypeID)
		
		If Request.Form("chkIsActive") = "on" Then
			blnIsActive = lhd_True
		Else
			blnIsActive = lhd_False
		End If

		dteLastUpdate = SQLDateFormat(Day(Now()), Month(Now()), Year(Now())) & " " & FormatDateTime(Now(), vbShortTime)
		intLastUpdateByID = intUserID

		
		' Check for required fields
		
		If Len(strCatName) = 0 Then
		
			Call DisplayError(1, "All required fields need to be entered, please go back and populate these fields")
			
		Else
		
			' Do nothing
			
		End If


		' Now save/update the Category details

		Set objCategory = New clsCategory

		objCategory.ID = intCatID

		If Not objCategory.Load Then
		
			' Category does not exist
'			Response.Write objCategory.LastError

		Else
		
			' Category exists
		
		End If

		' Check the the fields and leave Null if nothing is set.

		objCategory.CatName = strCatName
		objCategory.CatOrder = intCatOrder
		objCategory.CatDesc = strCatDesc
		objCategory.CaseTypeID = intCaseTypeID
		
		If Len(strCatDesc) > 0 Then 
			objCategory.CatDesc = strCatDesc
		End If
		
		objCategory.IsActive = blnIsActive
		objCategory.LastUpdate = dteLastUpdate
		objCategory.LastUpdateByID = intLastUpdateByID

						
		If Not objCategory.Update Then
						
			' Failed to create/save Category
							
		Else
						
			intCatID = objCategory.ID
			strHeading = Lang("Category_Saved")
						
		End If
						
		Set objCategory = Nothing
%>
		<TR>
		   <TD>
		      <TABLE class=Normal width="100%" border=0 cellSpacing=0 cellPadding=1>
				  <TR class="lhd_Heading1">
				    <TD colspan=5 align=center><%=strHeading%></TD>
				  </TR>
		      <TR>
		        <TD align=right colspan=5><a style="FONT-SIZE: 8pt" href="admCategoryList.asp?Page=1"><%=Lang("Manage_Categories")%></a></TD>
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
 				   <TD colspan=3 align=Left>Category changes have been successfully saved.</TD>
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
	
		' Mode: 1 - To create a new Category
		'		2 - Edit a Category
	
'		strMode = Request.QueryString("mode")
	
		If Request.QueryString.Count = 0 Then
	
			' Create a new record
	
			strMode = 1
			intCatID = 0
			strHeading = Lang("New_Category")
			
		Else
	
			' Edit a record determine by the Category ID passed via the QueryString
	
			strMode = 2
			intCatID = Request.QueryString("ID")
			strHeading = Lang("Modify_Category")
			
		End If
			

		Select Case strMode
		
			Case 1	' Create new Category
			
				strCatName = ""
				strCatDesc = ""
				intCatOrder = ""
				intCaseTypeID = 0

				blnIsActive = lhd_True
				strIsActiveHTML = "CHECKED"

				dteLastUpdate = ""
				intLastUpdateByID = 0


			Case 2  ' Edit Category

				' Get the Category ID we want to edit and load the record

				Set objCategory = New clsCategory

				objCategory.ID = intCatID
			
				If Not objCategory.Load Then
				
					' Couldn't load user for some reason
				
				Else
				
					strCatName = objCategory.CatName
					strCatDesc = objCategory.CatDesc
					intCatOrder = objCategory.CatOrder
					intCaseTypeID = objCategory.CaseTypeID
					
					If objCategory.IsActive = True Then
						strIsActiveHTML = "CHECKED"
					Else
						strIsActiveHTML = ""
					End If
					
					dteLastUpdate = objCategory.LastUpdate
					intLastUpdateByID = objCategory.LastUpdateByID

				End If

				Set objCategory = Nothing
				

			Case Else
				' Do nothing
				
		End Select

%>

  <TR>
    <TD>
	  <TABLE class=Normal cellSpacing=0 cellPadding=1 width="100%" border=0 bgColor=white>
	  <FORM action="admCategory.asp" method="post" id=frmCategory name=frmCategory>
	  <INPUT id=tbxCatID name=tbxCatID type=hidden value="<%=intCatID%>">
	  <INPUT id=tbxSave name=tbxSave type=hidden value="1">
	  
		<TR class="lhd_Heading1">
			<TD colspan=5 align=center><%=strHeading%></TD>
		</TR>
		<TR>
		  <TD align=right colspan=5><a style="FONT-SIZE: 8pt" href="admCategoryList.asp?Page=1"><%=Lang("Manage_Categories")%></a></TD>
		</TR>
		<TR>
		  <TD width="22%"><B><%=Lang("Case_Type")%>:</B></TD>
		  <TD width="25%">
				<SELECT id=cbxCaseType style="WIDTH: 100%" name=cbxCaseType>
				<%
				Set objCollection = New clsCollection
						
				objCollection.CollectionType = objCollection.clCaseType
				objCollection.Query = "SELECT CaseTypePK, CaseTypeName FROM tblCaseTypes ORDER BY CaseTypePK ASC"
						
				If Not objCollection.Load Then
							
				    Response.Write objCollection.LastError
							    
				Else
								
					If objCollection.BOF And objCollection.EOF Then
								
						' No records
								
					Else
							
						Do While Not objCollection.EOF
						
							If objCollection.Item.ID = intCaseTypeID THen
								%>
								<OPTION SELECTED VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.CaseTypeName%></OPTION>
								<%
							Else
								%>
								<OPTION VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.CaseTypeName%></OPTION>
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
		  <TD width="5%"></TD>
		  <TD width="18%"></TD>
		  <TD width="25%"></TD>
		</TR>
		<TR>
		  <TD><B><%=Lang("Category")%>:</B></TD>
		  <TD><INPUT id=tbxCatName name=tbxCatName style="WIDTH: 100%" value="<%=strCatName%>" ></TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		</TR>
		<TR>
		  <TD><%=Lang("Description")%>:</TD>
		  <TD colspan=4><INPUT id=tbxCatDesc name=tbxCatDesc style="WIDTH: 100%" value="<%=strCatDesc%>"></TD>
		</TR>
		<TR>
		  <TD><B><%=Lang("Order")%>:</B></TD>
		  <TD><INPUT id=tbxCatOrder name=tbxCatOrder style="WIDTH: 100%" value="<%=intCatOrder%>" ></TD>
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
  </TABLE></P></BODY></HTML>
  
<%
  
cnnDB.Close
Set cnnDB = Nothing
  
%>
