<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: admGroup.asp
'  Date:     $Date: 2004/03/18 06:03:18 $
'  Version:  $Revision: 1.6 $
'  Purpose:  Administration page for creating/modifing Groups
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


<SCRIPT language="VBScript" for="btnAdd" event="onClick">

  Set objMembers = Document.All("lbxMembers")
  Set objNonMembers = Document.All("lbxNonMembers")


  If objNonMembers.Options.SelectedIndex >= 0 Then

	' Add the members into the Members list

	Set eleOption = document.CreateElement("OPTION")
	eleOption.Value = objNonMembers.Options(objNonMembers.Options.SelectedIndex).Value
	eleOption.Text = objNonMembers.Options(objNonMembers.Options.SelectedIndex).Text

	objMembers.Add eleOption
	
	Set eleOption = Nothing


    ' The the Member ID to the MemberIDList box
    
	Set objMemberIDList = Document.All("tbxMemberIDList")
    objMemberIDList.Value = objMemberIDList.Value & objNonMembers.Options(objNonMembers.Options.SelectedIndex).Value & "#"
	Set objMemberIDList = Nothing


	' Remove the member from the Non-Members list

	objNonMembers.Options.Remove objNonMembers.Options.SelectedIndex

	

  Else

    ' Do nothing
    
  End If
  
  Set objMembers = Nothing
  Set objNonMembers = Nothing

</SCRIPT>

<SCRIPT language="VBScript" for="btnRemove" event="onClick">

  Set objMembers = Document.All("lbxMembers")
  Set objNonMembers = Document.All("lbxNonMembers")

  If objMembers.Options.SelectedIndex >= 0 Then

	' Add the members into the Members list

	Set eleOption = document.CreateElement("OPTION")
	eleOption.Value = objMembers.Options(objMembers.Options.SelectedIndex).Value
	eleOption.Text = objMembers.Options(objMembers.Options.SelectedIndex).Text

	objNonMembers.Add eleOption
	
	Set eleOption = Nothing


    ' Remove the Member ID to the MemberIDList box
    
	Set objMemberIDList = Document.All("tbxMemberIDList")
	objMemberIDList.Value = Replace( objMemberIDList.Value, "#" & objMembers.Options(objMembers.Options.SelectedIndex).Value & "#", "#")
	Set objMemberIDList = Nothing


	' Remove the member from the Non-Members list

	objMembers.Options.Remove objMembers.Options.SelectedIndex

  Else

    ' Do nothing
    
  End If
  
  Set objMembers = Nothing
  Set objNonMembers = Nothing

</SCRIPT>


</SCRIPT>


<LINK rel="stylesheet" type="text/css" href="Default.css">
<%
	Dim cnnDB
	Dim binUserPermMask
	Dim blnSave, blnIsActive
  Dim objGroup, objCollection
  Dim strGroupName, strGroupDesc, strMode, strIsActiveHTML
  Dim strMemberIDList, strHeading, strHTML
  Dim intUserID, intLastUpdateByID, intGroupID, intMemberID, intNextPos, intPrevPos
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

		strGroupName = Request.Form("tbxGroupName")
		strGroupDesc = Request.Form("tbxGroupDesc")
		intGroupID = Cint(Request.Form("tbxGroupID"))
    
    strMemberIDList = Request.Form("tbxMemberIDList")

		If Request.Form("chkIsActive") = "on" Then
			blnIsActive = lhd_True
		Else
			blnIsActive = lhd_False
		End If

		dteLastUpdate = SQLDateFormat(Day(Now()), Month(Now()), Year(Now())) & " " & FormatDateTime(Now(), vbShortTime)
		intLastUpdateByID = intUserID

		
		' Check for required fields
		
		If Len(strGroupName) = 0 Then
		
			Call DisplayError(1, "Group Name")
			
		Else
		
			' Do nothing
			
		End If


		' Now save/update the Groups details

		Set objGroup = New clsGroup

		objGroup.ID = intGroupID

		If Not objGroup.Load Then
		
			' Group does not exist, so we will be creating a new one.

		Else
		
			' Group exists, so we will be updating
		
		End If

		' Check the the fields and leave Null if nothing is set.

		objGroup.GroupName = strGroupName
		objGroup.GroupDesc = strGroupDesc
		objGroup.IsActive = blnIsActive
		objGroup.LastUpdate = dteLastUpdate
		objGroup.LastUpdateByID = intLastUpdateByID

						
		If Not objGroup.Update Then
						
			' Failed to create/save
			Response.Write objGroup.LastError & "<P>"
							
		Else
						
			intGroupID = objGroup.ID
			strHeading = Lang("Group_Saved")
						
		End If
		

		If objGroup.GroupMembers.BOF And objGroup.GroupMembers.EOF Then
		
		   ' No records found
		
		Else
		
		  While Not objGroup.GroupMembers.EOF
		  
		    If Instr(strMemberIDList, "#" & objGroup.GroupMembers.Item.ID & "#") > 0 Then
		    
		      ' Member ID found
		    
		    Else
		    
		      '  MemberID not in parsed Member ID List so effectively we need to remove the
		      ' member
		      
		      If objGroup.RemoveMember(objGroup.GroupMembers.Item.ID) = True Then
		        ' Removing member successful
		      Else
		        ' Removing member failed
		      End If
		      
		      ' Remove the ID from the parsed list
		      
		      Replace strMemberIDList,"#" & objGroup.GroupMembers.Item.ID & "#", "#"
		    
		    End IF
		    
		    objGroup.GroupMembers.MoveNext 
		  
		  WEnd
		
		End If

		
		' Now we add the remaining members list in the Member ID list to the group
		
		intPrevPos = 2
   	intNextPos = Instr(intPrevPos, strMemberIDList, "#")

		If intNextPos > 0 Then

		  Do		
		
		    If objGroup.AddMember( Mid(strMemberIDList, intPrevPos, intNextPos-intPrevPos) ) = True Then
		      ' Adding member successful
		    Else
		      ' Adding member failed
		    End If
		    
		    intPrevPos = intNextPos + 1
		    intNextPos = Instr(intPrevPos, strMemberIDList, "#")
		    
		  Loop Until intNextPos = 0
		  
		Else
		
		  ' Do nothing
		
		End If
		
						
		Set objGroup = Nothing
%>
		<TR>
		   <TD>
		      <TABLE class=Normal width="100%" border=0 cellSpacing=0 cellPadding=1>
				<TR class="lhd_Heading1">
				   <TD colspan=5 align=center><%=strHeading%></TD>
				</TR>
		      <TR>
		        <TD align=right colspan=5><a style="FONT-SIZE: 8pt" href="admGroupList.asp?Page=1"><%=Lang("Manage_Groups")%></a></TD>
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
 				   <TD colspan=3 align=left>Group information has been successfully saved.</TD>
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
	
		' Mode: 1 - To create a new department
		'		2 - Edit a department
	
'		strMode = Request.QueryString("mode")
	
		If Request.QueryString.Count = 0 Then
	
			' Create a new record
	
			strMode = 1
			intGroupID = 0
			strHeading = Lang("New_Group")
			
		Else
	
			' Edit a record determine by the DeptID passed via the QueryString
	
			strMode = 2
			intGroupID = CInt(Request.QueryString("ID"))
			strHeading = Lang("Modify_Group")
			
		End If
			

		Select Case strMode
		
			Case 1	' Create new Group
			
				strGroupName = ""
				strGroupDesc = ""

				blnIsActive = lhd_True
				strIsActiveHTML = "CHECKED"

				dteLastUpdate = ""
				intLastUpdateByID = 0


			Case 2  ' Edit Group

				' Get the Group ID we want to edit and load the record

				Set objGroup = New clsGroup

				objGroup.ID = intGroupID
			
				If Not objGroup.Load Then
				
					' Couldn't load Group
					Response.Write objGroup.LastError
				
				Else
				
					strGroupName = objGroup.GroupName
					strGroupDesc = objGroup.GroupDesc
					
					If objGroup.IsActive = True Then
						strIsActiveHTML = "CHECKED"
					Else
						strIsActiveHTML = ""
					End If
					
					dteLastUpdate = objGroup.LastUpdate
					intLastUpdateByID = objGroup.LastUpdateByID

				End If

				Set objGroup = Nothing
				

			Case Else
				' Do nothing
				
		End Select


		' Build members list

		Set objCollection = New clsCollection
		      
		objCollection.CollectionType = objCollection.clContact
		objCollection.Query = "SELECT tblContacts.* FROM tblContacts INNER JOIN tblGroupMembers ON tblContacts.ContactPK = tblGroupMembers.ContactFK WHERE tblGroupMembers.GroupFK=" & intGroupID & " ORDER BY tblContacts.Username ASC"

		If Not objCollection.Load Then
		      
		  ' Collection didn't load
		        
		Else
		      
		  If objCollection.BOF And objCollection.EOF Then
		        
		    ' No records returned
				  
		  Else
				
   	    strHTML = ""
     		      
		    Do While Not objCollection.EOF
				  
		      strHTML = strHTML & "<OPTION value=" & objCollection.Item.ID & "> " & objCollection.Item.Username & " (" & objCollection.Item.FName & " " & objCollection.Item.LName & ")</OPTION>"

		      strMemberIDList = strMemberIDList & objCollection.Item.ID & "#"

 		  	  objCollection.MoveNext
					
		    Loop
				
		  End If
		      
		End If
		    
		Set objCollection = Nothing
%>

  <TR>
    <TD>
	  
	  <TABLE class=Normal cellSpacing=0 cellPadding=1 width="100%" border=0 bgColor=white>
	  <FORM action="admGroup.asp" method="post" id=frmGroup name=frmGroup>
	  <INPUT id=tbxGroupID name=tbxGroupID type=hidden value="<%=intGroupID%>">
	  <INPUT id=tbxSave name=tbxSave type=hidden value="1">
	  <INPUT id=tbxMemberIDList name=tbxMemberIDList type=hidden value="#<%=strMemberIDList%>">
		<TR class="lhd_Heading1">
			<TD colspan=5 align=center><%=strHeading%></TD>
		</TR>
		<TR>
		  <TD align=right colspan=5><a style="FONT-SIZE: 8pt" href="admGroupList.asp?Page=1"><%=Lang("Manage_Groups")%></a></TD>
		</TR>
		<TR>
		  <TD width="22%"><B><%=Lang("Group_Name")%>:</B></TD>
		  <TD width="25%"><INPUT id=tbxGroupName name=tbxGroupName style="WIDTH: 100%" value="<%=strGroupName%>" ></TD>
		  <TD width="5%"></TD>
		  <TD width="18%"></TD>
		  <TD width="25%"></TD>
		</TR>
		<TR>
		  <TD><%=Lang("Description")%>:</TD>
		  <TD colspan=4><INPUT id=tbxGroupDesc name=tbxGroupDesc style="WIDTH: 100%" value="<%=strGroupDesc%>"></TD>
		</TR>
		<TR>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		</TR>
		<TR>
		  <TD valign=top><%=Lang("Members")%>:</TD>
		  <TD>
		    <SELECT style="WIDTH: 100%" id=lbxMembers size=7 name=lbxMembers> 
			<%
			  Response.Write strHTML
	        %>
            </SELECT>
          </TD>
		  <TD align=center>
		  </TD>
		  <TD valign=top><%=Lang("Non_Members")%>:</TD>
		  <TD>
		    <SELECT style="WIDTH: 100%" id=lbxNonMembers size=7 name=lbxNonMembers> 
		      <%
		      Set objCollection = New clsCollection
		      
		      objCollection.CollectionType = objCollection.clContact
		      objCollection.Query = "Select * FROM tblContacts " & _
					    "Where ContactPK Not IN   " & _
						"(Select ContactPK  FROM tblContacts " & _
						"Inner JOIN tblGroupMembers ON tblContacts.ContactPK = tblGroupMembers.ContactFK " & _
						"WHERE tblGroupMembers.GroupFK = " & intGroupID & " Or tblGroupMembers.GroupFK Is Null ) " & _
					     "ORDER BY tblContacts.Username ASC"
		      
		      If Not objCollection.Load Then
		      
    				' Collection didn't load
		        
		      Else
		      
		        If objCollection.BOF And objCollection.EOF Then
		        
		  		    ' No records returned
				  
			    	Else
				
     		      strHTML = ""
     		      
				      Do While Not objCollection.EOF
				  
				        strHTML = strHTML & "<OPTION value=" & objCollection.Item.ID & "> " & objCollection.Item.Username & " (" & objCollection.Item.FName & " " & objCollection.Item.LName & ")</OPTION>"

				        objCollection.MoveNext

				      Loop
				
		          Response.Write strHTML
				
    				End If
		      
		      End If
		    
		      Set objCollection = Nothing
		      %>
            </SELECT>
          </TD>
		</TR>
		<TR>
		  <TD></TD>
		  <TD></TD>
		  <TD colspan=2 align=center>
		    <INPUT style="FONT-SIZE: xx-small; WIDTH=50px" id=btnAdd name=btnAdd type="button" value="<%=Lang("Add")%>">&nbsp;&nbsp;
		    <INPUT style="FONT-SIZE: xx-small; WIDTH=50px" id=btnRemove name=btnRemove type="button" value="<%=Lang("Remove")%>">
		  </TD>
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
