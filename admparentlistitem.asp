<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: admParentListItem.asp
'  Date:     $Date: 2004/03/10 08:21:12 $
'  Version:  $Revision: 1.5 $
'  Purpose:  Administration page for creating/modifing Parent Lists
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
    Dim objParentListItem, objCollection
    Dim strParentListItemName, strMode, strIsActiveHTML, strHeading
    Dim intUserID, intLastUpdateByID, intParentListItemID, intParentListItemOrder
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

		intParentListItemID = Cint(Request.Form("tbxParentListItemID"))
		strParentListItemName = Request.Form("tbxParentListItemName")
		intParentListItemOrder = CInt(Request.Form("tbxParentListItemOrder"))

		If Request.Form("chkIsActive") = "on" Then
			blnIsActive = lhd_True
		Else
			blnIsActive = lhd_False
		End If

		dteLastUpdate = SQLDateFormat(Day(Now()), Month(Now()), Year(Now())) & " " & FormatDateTime(Now(), vbShortTime)
		intLastUpdateByID = intUserID

		
		' Check for required fields
		
		If Len(strParentListItemName) = 0 Then
		
			Call DisplayError(1, "All required fields need to be entered, please go back and populate these fields")
			
		Else
		
			' Do nothing
			
		End If


		' Now save/update the Parent List Item details

		Set objParentListItem = New clsListItem

		objParentListItem.ID = intParentListItemID

		If Not objParentListItem.Load Then
		
			' Parent List Item does not exist
			Response.Write objParentListItem.LastError & "<P>"

		Else
		
			' Parent List Item exists
		
		End If


		' Check the the fields and leave Null if nothing is set.

		objParentListItem.ItemName = strParentListItemName
		objParentListItem.ItemOrder = intParentListItemOrder
		objParentListItem.IsActive = blnIsActive
		objParentListItem.LastUpdate = dteLastUpdate
		objParentListItem.LastUpdateByID = intLastUpdateByID
						
		If Not objParentListItem.Update Then
						
			' Failed to create/save department
			Response.Write objParentListItem.LastError & "<P>"
							
		Else
						
			intParentListItemID = objParentListItem.ID
			strHeading = Lang("List_Saved")
						
		End If
						
		Set objParentListItem = Nothing
%>
		<TR>
		   <TD>
		      <TABLE class=Normal cellSpacing=0 cellPadding=1 width="100%" border=0 bgColor=white>
		         <TR class="lhd_Heading1" >
 				 <TD colspan=5 align=middle><%=strHeading%></TD>
				 </TR>
		      <TR>
		        <TD align=right colspan=5>
		          <a style="FONT-SIZE: 8pt" href="admParentListItemList.asp?Page=1"><%=Lang("Manage_Lists")%></a>&nbsp;|&nbsp;
		          <a style="FONT-SIZE: 8pt" href="admListItemList.asp?List=<%=intParentListItemID%>&Page=1"><%=Lang("Manage_List_Items")%></a>
		        </TD>
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
 				 <TD colspan=3 align=left>List information has been successfully saved.</TD>
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
	
		' Mode: 1 - To create a new Parent List Item
		'		2 - Edit a Parent List Item
	
'		strMode = Request.QueryString("mode")
	
		If Request.QueryString.Count = 0 Then
	
			' Create a new record
	
			strMode = 1
			intParentListItemID = 0
			strHeading = Lang("New_List")
			
		Else
	
			' Edit a record determine by the Parent List Item ID passed via the QueryString
	
			strMode = 2
			intParentListItemID = Request.QueryString("id")
			strHeading = Lang("Modify_List")
			
		End If
			

		Select Case strMode
		
			Case 1	' Create new List Item
			
				strParentListItemName = ""
				intParentListItemOrder = 0
				blnIsActive = lhd_True
				dteLastUpdate = ""
				intLastUpdateByID = 0

				strIsActiveHTML = "CHECKED"


			Case 2  ' Edit List Item

				' Get the Parent List Item ID we want to edit and load the record

				Set objParentListItem = New clsListItem

				objParentListItem.ID = intParentListItemID
			
				If Not objParentListItem.Load Then
				
					' Couldn't load Parent List Item for some reason
					Response.Write objParentListItem.LastError & "<P>"
				
				Else
				
					strParentListItemName = objParentListItem.ItemName
					intParentListItemOrder = objParentListItem.ItemOrder
					
					If objParentListItem.IsActive = True Then
						strIsActiveHTML = "CHECKED"
					Else
						strIsActiveHTML = ""
					End If
					
					dteLastUpdate = objParentListItem.LastUpdate
					intLastUpdateByID = objParentListItem.LastUpdateByID

				End If

				Set objParentListItem = Nothing
				

			Case Else
				' Do nothing
				
		End Select

%>

  <TR>
    <TD>
	  <TABLE class=Normal cellSpacing=0 cellPadding=1 width="100%" border=0 bgColor=white>
	  <FORM action="admParentListItem.asp" method="post" id=frmParentListItem name=frmParentListItem>
	  <INPUT id=tbxParentListItemID name=tbxParentListItemID type=hidden value="<%=intParentListItemID%>">
	  <INPUT id=tbxSave name=tbxSave type=hidden value="1">
	  
		<TR class="lhd_Heading1">
			<TD colspan=5 align=center><%=strHeading%></TD>
		</TR>
		<TR>
		  <TD align=right colspan=5>
		    <a style="FONT-SIZE: 8pt" href="admParentListItemList.asp?Page=1"><%=Lang("Manage_Lists")%></a>&nbsp;|&nbsp;
		    <a style="FONT-SIZE: 8pt" href="admListItemList.asp?List=<%=intParentListItemID%>&Page=1"><%=Lang("Manage_List_Items")%></a>
		  </TD>
		</TR>
		<TR>
		  <TD width="22%"><B><%=Lang("List_Name")%>:</B></TD>
		  <TD width="25%"><INPUT id=tbxParentListItemName name=tbxParentListItemName style="WIDTH: 100%" value="<%=strParentListItemName%>" ></TD>
		  <TD width="5%"></TD>
		  <TD width="18%"></TD>
		  <TD width="25%"></TD>
		</TR>
		<TR>
		  <TD><%=Lang("Item_Order")%>:</TD>
		  <TD><INPUT id=tbxParentListItemOrder name=tbxParentListItemOrder
            style="WIDTH: 100%" value="<%=intParentListItemOrder%>"></TD>
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
		  <TD colspan=3></TD>
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
