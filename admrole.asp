<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: admRole.asp
'  Date:     $Date: 2004/03/10 08:21:12 $
'  Version:  $Revision: 1.5 $
'  Purpose:  Administration page for creating/modifing Roles & Permissions
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
	Dim binUserPermMask, binRoleMask, binRequiredPerm
	Dim blnSave, blnIsActive
  Dim objRole, objCollection
  Dim strRoleName, strRoleDesc, strMode, strIsActiveHTML, strHTML, strHeading
  Dim intUserID, intLastUpdateByID, intRoleID
  Dim dteLastUpdate
  Dim rsPerm
  Dim lngPermTotal


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
<TABLE align=center cellSpacing=1 cellPadding=1 width="680" border=0>
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

		strRoleName = Request.Form("tbxRoleName")
		strRoleDesc = Request.Form("tbxRoleDesc")
		
		' Build the list of permissions

		Set rsPerm = Server.CreateObject("ADODB.Recordset")
				
		rsPerm.Open "SELECT PermLabel FROM tblPermissions ORDER BY PermLabel ASC", cnnDB

		If rsPerm.BOF And rsPerm.EOF Then
				
			' No Records
					
		Else
		
			binRoleMask = 0

			Do While Not rsPerm.EOF
				If Request.Form("chk" &  rsPerm("PermLabel")) = "on" Then
					binRoleMask = binRoleMask + Clng(Request.Form( "tbx" & rsPerm("PermLabel") ))
				Else
					' Do nothing
				End If
				rsPerm.MoveNext 
			Loop
				
		End If
				
		rsPerm.Close
		Set rsPerm = Nothing
		
		binRoleMask = "&H" & Right("0000000000000000" & Hex(binRoleMask), 16)

		intRoleID = Cint(Request.Form("tbxRoleID"))

		If Request.Form("chkIsActive") = "on" Then
			blnIsActive = lhd_True
		Else
			blnIsActive = lhd_False
		End If

		dteLastUpdate = SQLDateFormat(Day(Now()), Month(Now()), Year(Now())) & " " & FormatDateTime(Now(), vbShortTime)
		intLastUpdateByID = intUserID

		
		' Check for required fields
		
		If Len(strRoleName) = 0 Then
		
			Call DisplayError(1, "All required fields need to be entered, please go back and populate these fields")
			
		Else
		
			' Do nothing
			
		End If


		' Now save/update the Role details

		Set objRole = New clsRole

		objRole.ID = intRoleID

		If Not objRole.Load Then
		
			' Role does not exist
			Response.Write objRole.LastError

		Else
		
			' Role exists
		
		End If

		' Check the the fields and leave Null if nothing is set.

		objRole.RoleName = strRoleName
		objRole.RoleDesc = strRoleDesc
		objRole.RoleMask = binRoleMask
		objRole.IsActive = blnIsActive
		objRole.LastUpdate = dteLastUpdate
		objRole.LastUpdateByID = intLastUpdateByID

						
		If Not objRole.Update Then
						
			' Failed to create/save Role
			Response.Write objRole.LastError
							
		Else
						
			intRoleID = objRole.ID
			strHeading = Lang("Role") & " " & Lang("Saved")
						
		End If
						
		Set objRole = Nothing
%>
		<TR>
		   <TD>
		      <TABLE class=Normal cellSpacing=0 cellPadding=1 width="100%" border=0 bgColor=white>
		         <TR class="lhd_Heading1" >
 				 <TD colspan=5 align=middle><%=strHeading%></TD>
				 </TR>
		      <TR>
		        <TD align=right colspan=5><a style="FONT-SIZE: 8pt" href="admRoleList.asp?Page=1"><%=Lang("Manage_Roles")%></a></TD>
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
 				 <TD colspan=3 align=left>Role information has been successfully saved.</TD>
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
	
		' Mode: 1 - To create a new Role
		'		2 - Edit a Role
	
'		strMode = Request.QueryString("mode")
	
		If Request.QueryString.Count = 0 Then
	
			' Create a new record
	
			strMode = 1
			intRoleID = 0
			strHeading = Lang("New") & " " & Lang("Role")
			
		Else
	
			' Edit a record determine by the Role ID passed via the QueryString
	
			strMode = 2
			intRoleID = Request.QueryString("id")
			strHeading = Lang("Modify") & " " & Lang("Role")
			
		End If
			

		Select Case strMode
		
			Case 1	' Create new role
			
				strRoleName = ""
				strRoleDesc = ""
				binRoleMask = &H0000000000000000

				blnIsActive = lhd_True
				strIsActiveHTML = "CHECKED"

				dteLastUpdate = ""
				intLastUpdateByID = 0


			Case 2  ' Edit Role

				' Get the Role ID we want to edit and load the record

				Set objRole = New clsRole

				objRole.ID = intRoleID
			
				If Not objRole.Load Then
				
					' Couldn't load user for some reason
					Response.Write objRole.LastError & "<P>"
				
				Else
				
					strRoleName = objRole.RoleName
					strRoleDesc = objRole.RoleDesc
					binRoleMask = objRole.RoleMask
					
					If objRole.IsActive = True Then
						strIsActiveHTML = "CHECKED"
					Else
						strIsActiveHTML = ""
					End If
					
					dteLastUpdate = objRole.LastUpdate
					intLastUpdateByID = objRole.LastUpdateByID

				End If

				Set objRole = Nothing
				

			Case Else
				' Do nothing
				
		End Select


		' Build the list of permissions

		Set rsPerm = Server.CreateObject("ADODB.Recordset")
				
'		rsPerm.Open "SELECT PermLabel, PermDesc, CONVERT(INT, PermByte) AS 'PermByte' FROM tblPermissions", cnnDB
		rsPerm.Open "SELECT PermLabel, PermDesc, PermByte FROM tblPermissions ORDER BY PermLabel ASC", cnnDB

		If rsPerm.BOF And rsPerm.EOF Then
				
			' No Records
					
		Else

			strHTML = ""

			strHTML = strHTML & "<TR>" & Chr(13)
			strHTML = strHTML & "<TD>" & Lang("Permissions") & ":</TD>" & Chr(13)
				If rsPerm("PermByte") <> (rsPerm("PermByte") And binRoleMask) Then
					strHTML = strHTML & "<TD><INPUT id=chk" & rsPerm("PermLabel") & " name=chk" & rsPerm("PermLabel") & " type=checkbox></TD>" & Chr(13)
				Else
					strHTML = strHTML & "<TD><INPUT id=chk" & rsPerm("PermLabel") & " name=chk" & rsPerm("PermLabel") & " type=checkbox checked></TD>" & Chr(13)
				End If
				strHTML = strHTML & "<TD colspan=3><INPUT id=tbx" & rsPerm("PermLabel") & " name=tbx" & rsPerm("PermLabel") & " value=" & rsPerm("PermByte") & " type=hidden>" & rsPerm("PermDesc") & "</TD>" & Chr(13)
			strHTML = strHTML & "</TR>" & Chr(13)

			rsPerm.MoveNext 

			Do While Not rsPerm.EOF
				
				strHTML = strHTML & "<TR>" & Chr(13)
				strHTML = strHTML & "<TD></TD>" & Chr(13)
				If rsPerm("PermByte") <> (rsPerm("PermByte") And binRoleMask) Then
					strHTML = strHTML & "<TD><INPUT id=chk" & rsPerm("PermLabel") & " name=chk" & rsPerm("PermLabel") & " type=checkbox></TD>" & Chr(13)
				Else
					strHTML = strHTML & "<TD><INPUT id=chk" & rsPerm("PermLabel") & " name=chk" & rsPerm("PermLabel") & " type=checkbox checked></TD>" & Chr(13)
				End If
				strHTML = strHTML & "<TD colspan=3><INPUT id=tbx" & rsPerm("PermLabel") & " name=tbx" & rsPerm("PermLabel") & " value=" & rsPerm("PermByte") & " type=hidden>" & rsPerm("PermDesc") & "</TD>" & Chr(13)
				strHTML = strHTML & "</TR>" & Chr(13)

				rsPerm.MoveNext 
					
			Loop
				
		End If
				
		rsPerm.Close
		Set rsPerm = Nothing


%>

  <TR>
    <TD>
	  <TABLE class=Normal cellSpacing=0 cellPadding=1 width="100%" border=0 bgColor=white>
	  <FORM action="admRole.asp" method="post" id=frmRole name=frmRole>
	  <INPUT id=tbxRoleID name=tbxRoleID type=hidden value="<%=intRoleID%>">
	  <INPUT id=tbxSave name=tbxSave type=hidden value="1">

		<TR class="lhd_Heading1">
			<TD colspan=5 align=center><%=strHeading%></TD>
		</TR>
		<TR>
		  <TD align=right colspan=5><a style="FONT-SIZE: 8pt" href="admRoleList.asp?Page=1"><%=Lang("Manage_Roles")%></a></TD>
		</TR>
		<TR>
		  <TD width="20%"><B><%=Lang("Role_Name")%>:</B></TD>
		  <TD width="5%"><INPUT id=tbxRoleName name=tbxRoleName style="WIDTH: 100%" value="<%=strRoleName%>" ></TD>
		  <TD width="25%"></TD>
		  <TD width="25%"></TD>
		  <TD width="25%" align=right></TD>
		</TR>
		<TR>
		  <TD><%=Lang("Description")%>:</TD>
		  <TD colspan=4><INPUT id=tbxRoleDesc name=tbxRoleDesc 
            style="WIDTH: 100%" value="<%=strRoleDesc%>"></TD>
		</TR>
		<TR>
		  <TD></TD>
		  <TD align="Right" colspan="4"><INPUT id=chkIsActive name=chkIsActive type=checkbox <%=strIsActiveHTML%>>&nbsp;&nbsp;<%=Lang("Is_Active")%></TD>
		</TR>
		<TR>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		</TR>
		<%
		Response.Write strHTML	
		%>
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
  </TABLE></P></BODY>
  
</HTML>

<%
cnnDB.Close
Set cnnDB = Nothing
%>
