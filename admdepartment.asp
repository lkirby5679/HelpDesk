<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: admDepartment.asp
'  Date:     $Date: 2004/03/10 08:21:12 $
'  Version:  $Revision: 1.5 $
'  Purpose:  Administration page for creating/modifing Departments
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
    <META content="MSHTML 6.00.2600.0" name="GENERATOR">
  </HEAD>
  <LINK rel="stylesheet" type="text/css" href="Default.css">
    <%
	Dim cnnDB
	Dim binUserPermMask, binRequiredPerm
	Dim blnSave, blnIsActive
  Dim objDept, objCollection
  Dim strDeptName, strDeptDesc, strMode, strIsActiveHTML, strHeading
  Dim intUserID, intLastUpdateByID, intDeptID, intOrgID
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
      <P align="center">
        <TABLE class="Normal" align="center" cellSpacing="1" cellPadding="1" width="680" border="0">
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

		strDeptName = Request.Form("tbxDeptName")
		strDeptDesc = Request.Form("tbxDeptDesc")
		intOrgID = Cint(Request.Form("cbxOrganisation"))
		intDeptID = Cint(Request.Form("tbxDeptID"))

		If Request.Form("chkIsActive") = "on" Then
			blnIsActive = lhd_True
		Else
			blnIsActive = lhd_False
		End If

		dteLastUpdate = SQLDateFormat(Day(Now()), Month(Now()), Year(Now())) & " " & FormatDateTime(Now(), vbShortTime)
		intLastUpdateByID = intUserID

		
		' Check for required fields
		
		If Len(strDeptName) = 0 Or intOrgID = 0 Then
		
			Call DisplayError(1, "All required fields need to be entered, please go back and populate these fields")
			
		Else
		
			' Do nothing
			
		End If


		' Now save/update the Department details

		Set objDept = New clsDepartment

		objDept.ID = intDeptID

		If Not objDept.Load Then
		
			' Department does not exist, so we are set to create a new one

		Else
		
			' Department exists, so we update
		
		End If

		' Check the the fields and leave Null if nothing is set.

		objDept.OrgID = intOrgID
		objDept.DeptName = strDeptName
		
		If Len(strDeptDesc) > 0 Then 
			objDept.DeptDesc = strDeptDesc
		End If
		
		objDept.IsActive = blnIsActive
		objDept.LastUpdate = dteLastUpdate
		objDept.LastUpdateByID = intLastUpdateByID

						
		If Not objDept.Update Then
						
			' Failed to create/save department
			Response.Write objDept.LastError
							
		Else
						
			intDeptID = objDept.ID
			strHeading = Lang("Department_Saved")
						
		End If
						
		Set objDept = Nothing
%>
            <TR>
              <TD>
                <TABLE class="Normal" cellSpacing="0" cellPadding="1" width="100%" border="0" bgColor="white">
                  <TR class="lhd_Heading1">
                    <TD></TD>
                    <TD colspan="3" align="middle"><%=strHeading%></TD>
                    <TD></TD>
                  </TR>
                  <TR>
                    <TD align="right" colspan="5"><a style="FONT-SIZE: 8pt" href="admDepartmentList.asp?Page=1"><%=Lang("Manage_Departments")%></a></TD>
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
                    <TD colspan="3" align="left">Department information has been successfully saved.</TD>
                    <TD></TD>
                  </TR>
                  <TR>
                    <TD></TD>
                    <TD colspan="3"></TD>
                    <TD></TD>
                  </TR>
                </TABLE>
              </TD>
            </TR>
            <%
	Else
	
		' Mode: 1 - To create a new department
		'		2 - Edit a department
	
		If Request.QueryString.Count = 0 Then
	
			' Create a new record
	
			strMode = 1
			intDeptID = 0
			strHeading = Lang("New_Department")
			
		Else
	
			' Edit a record determine by the DeptID passed via the QueryString
	
			strMode = 2
			intDeptID = Request.QueryString("ID")
			strHeading = Lang("Modify_Department")
			
		End If
			

		Select Case strMode
		
			Case 1	' Create new department
			
				strDeptName = ""
				strDeptDesc = ""
				intOrgID = 0

				blnIsActive = lhd_True
				strIsActiveHTML = "CHECKED"

				dteLastUpdate = ""
				intLastUpdateByID = 0


			Case 2  ' Edit department

				' Get the Department ID we want to edit and load the record

				Set objDept = New clsDepartment

				objDept.ID = intDeptID
			
				If Not objDept.Load Then
				
					' Couldn't load user for some reason
					Response.Write objDept.LastError
				
				Else
				
					strDeptName = objDept.DeptName
					strDeptDesc = objDept.DeptDesc
					intOrgID = objDept.OrgID
					
					If objDept.IsActive = True Then
						strIsActiveHTML = "CHECKED"
					Else
						strIsActiveHTML = ""
					End If
					
					dteLastUpdate = objDept.LastUpdate
					intLastUpdateByID = objDept.LastUpdateByID

				End If

				Set objDept = Nothing
				

			Case Else
				' Do nothing
				
		End Select

%>
            <TR>
              <TD>
                <TABLE class="Normal" cellSpacing="0" cellPadding="1" width="100%" border="0" bgColor="white">
                  <FORM action="admDepartment.asp" method="post" id="frmDept" name="frmDept">
                    <INPUT id=tbxDeptID name=tbxDeptID type=hidden value="<%=intDeptID%>"> <INPUT id="tbxSave" name="tbxSave" type="hidden" value="1">
                    <TR class="lhd_Heading1">
                      <TD colspan="5" align="center"><%=strHeading%></TD>
                    </TR>
                    <TR>
                      <TD align="right" colspan="5"><a style="FONT-SIZE: 8pt" href="admDepartmentList.asp?Page=1"><%=Lang("Manage_Departments")%></a></TD>
                    </TR>
                    <TR>
                      <TD><B><%=Lang("Organisation")%>:<B></TD>
                      <TD>
                        <SELECT id="cbxOrganisation" name="cbxOrganisation" style="WIDTH: 100%">
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
                      <TD></TD>
                    </TR>
                    <TR>
                      <TD width="22%"><B><%=Lang("Department")%>:</B></TD>
                      <TD width="25%"><INPUT id=tbxDeptName name=tbxDeptName style="WIDTH: 100%" value="<%=strDeptName%>" ></TD>
                      <TD width="5%"></TD>
                      <TD width="18%"></TD>
                      <TD width="25%"></TD>
                    </TR>
                    <TR>
                      <TD><%=Lang("Description")%>:</TD>
                      <TD colspan="4"><INPUT id=tbxDeptDesc name=tbxDeptDesc             style="WIDTH: 100%" value="<%=strDeptDesc%>"></TD>
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
                      <TD align="right"><INPUT style="WIDTH: 80px; BACKGROUND-COLOR: white" id=btnSave name=btnSave type=submit value="<%=Lang("Save")%>"></TD>
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
    </BODY></HTML>
<%
  
cnnDB.Close
Set cnnDB = Nothing
  
%>
