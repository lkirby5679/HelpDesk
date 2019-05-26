<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: admListItem.asp
'  Date:     $Date: 2004/03/10 08:21:12 $
'  Version:  $Revision: 1.6 $
'  Purpose:  Administration page for creating/modifing List Items
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
  Dim objListItem, objParentListItem
  Dim strListItemName, strMode, strIsActiveHTML, strParentListItem, strHeading
  Dim intUserID, intLastUpdateByID, intListItemID, intParentListItemID, intListItemOrder
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
<TABLE align=center cellSpacing=1 cellPadding=1 width="680" border=0>
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

		intListItemID = Cint(Request.Form("tbxListItemID"))
		strListItemName = Request.Form("tbxListItemName")
		intParentListItemID = CInt(Request.Form("tbxParentListItemID"))
		intListItemOrder = CInt(Request.Form("tbxListItemOrder"))

		If Request.Form("chkIsActive") = "on" Then
			blnIsActive = lhd_True
		Else
			blnIsActive = lhd_False
		End If

		dteLastUpdate = SQLDateFormat(Day(Now()), Month(Now()), Year(Now())) & " " & FormatDateTime(Now(), vbShortTime)
		intLastUpdateByID = intUserID

		
		' Check for required fields
		
		If Len(strListItemName) = 0 Then
		
			Call DisplayError(1, "All required fields need to be entered, please go back and populate these fields")
			
		Else
		
			' Do nothing
			
		End If


		' Now save/update the List Item details

		Set objListItem = New clsListItem

		objListItem.ID = intListItemID

		If Not objListItem.Load Then
		
			' List Item does not exist
			Response.Write objListItem.LastError

		Else
		
			' List Item exists
		
		End If


		' Check the the fields and leave Null if nothing is set.

		objListItem.ParentListItemID = intParentListItemID
		objListItem.ItemName = strListItemName
		objListItem.ItemOrder = intListItemOrder
		objListItem.IsActive = blnIsActive
		objListItem.LastUpdate = CDate(dteLastUpdate)
		objListItem.LastUpdateByID = intLastUpdateByID
						
		If Not objListItem.Update Then
						
			' Failed to create/save List Item
			Response.Write objListItem.LastError
							
		Else
						
			intListItemID = objListItem.ID
			strHeading = Lang("List_Item_Saved")
						
		End If
						
		Set objListItem = Nothing
%>
		<TR>
		   <TD>
		      <TABLE class=Normal cellSpacing=0 cellPadding=1 width="100%" border=0 bgColor=white>
		         <TR>
					<TD></TD>
 					<TD class="lhd_Heading1" colspan=3 align=middle><%=strHeading%></TD>
					<TD></TD>
				 </TR>
		      <TR>
		        <TD align=right colspan=5><a style="FONT-SIZE: 8pt" href="admListItemList.asp?List=<%=intParentListItemID%>&Page=1"><%=Lang("Manage_List_Items")%></a></TD>
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
 					<TD colspan=3 align=left>List Item information has been successfully saved.</TD>
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
	
		' Mode: 1 - To create a new List Item
		'		2 - Edit a List Item
	
'		strMode = Request.QueryString("mode")
	
		intParentListItemID = Request.QueryString("List")
	
		If intParentListItemID > 0 Then
	
			' Create a new record
	
			strMode = 1
			intListItemID = 0
			
			Set objParentListItem = New clsListItem
		
			objParentListItem.ID = intParentListItemID
		
			If Not objParentListItem.Load Then
				' Item not loaded
			Else
				strParentListItem = objParentListItem.ItemName
			End If
		
			Set objParentListItem = Nothing
			
		Else
	
			' Edit a record determine by the List Item ID passed via the QueryString
	
			strMode = 2
			intListItemID = Request.QueryString("ID")
			
		End If
			
		

		Select Case strMode
		
			Case 1	' Create new List Item
			
				strListItemName = ""
				intListItemOrder = 0
				blnIsActive = lhd_True
				dteLastUpdate = ""
				intLastUpdateByID = 0

				strIsActiveHTML = "CHECKED"
				
				strHeading = "New " & strParentListItem & " Item"

			Case 2  ' Edit List Item

				' Get the Department ID we want to edit and load the record

				Set objListItem = New clsListItem

				objListItem.ID = intListItemID
			
				If Not objListItem.Load Then
				
					' Couldn't load user for some reason
					Response.Write objListItem.LastError
				
				Else
				
					strListItemName = objListItem.ItemName
					intParentListItemID = objListItem.ParentListItemID
					intListItemOrder = objListItem.ItemOrder
					
					If objListItem.IsActive = True Then
						strIsActiveHTML = "CHECKED"
					Else
						strIsActiveHTML = ""
					End If
					
					dteLastUpdate = objListItem.LastUpdate
					intLastUpdateByID = objListItem.LastUpdateByID

					strParentListItem = objListItem.ParentListItem.ItemName

				End If

				Set objListItem = Nothing
				
				strHeading = "Modify " & strParentListItem & " Item"

			Case Else
				' Do nothing
				
		End Select

%>

  <TR>
    <TD>
	  <TABLE class=Normal cellSpacing=0 cellPadding=1 width="100%" border=0 bgColor=white>
	  <FORM action="admListItem.asp" method="post" id=fmmListItem name=frmListItem>
	  <INPUT id=tbxListItemID name=tbxListItemID type=hidden value="<%=intListItemID%>">
	  <INPUT id=tbxParentListItemID name=tbxParentListItemID type=hidden value="<%=intParentListItemID%>">
	  <INPUT id=tbxSave name=tbxSave type=hidden value="1">
	  
		<TR class="lhd_Heading1">
			<TD colspan=5 align=center><%=strHeading%></TD>
		</TR>
		<TR>
		  <TD align=right colspan=5><a style="FONT-SIZE: 8pt" href="admListItemList.asp?List=<%=intParentListItemID%>&Page=1"><%=Lang("Manage_List_Items")%></a></TD>
		</TR>
		<TR>
		<TD><%=Lang("List_Name")%>:</TD>
		<TD><%=strParentListItem%></TD>
		<TD></TD>
		<TD></TD>
	    <TD></TD></TR>
		<TR>
		  <TD width="22%"><B><%=Lang("Item_Name")%>:</B></TD>
		  <TD width="25%"><INPUT id=tbxListItemName name=tbxListItemName style="WIDTH: 100%" value="<%=strListItemName%>" ></TD>
		  <TD width="5%"></TD>
		  <TD width="18%"></TD>
		  <TD width="25%"></TD>
		</TR>
		<TR>
		  <TD><B><%=Lang("Item_Order")%>:</B></TD>
		  <TD><INPUT id=tbxListItemOrder name=tbxListItemOrder
            style="WIDTH: 100%" value="<%=intListItemOrder%>"></TD>
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
</P></BODY></HTML>

<%
cnnDB.Close
Set cnnDB = Nothing
%>
