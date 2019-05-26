<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: caseNew.asp
'  Date:     $Date: 2004/03/25 03:13:48 $
'  Version:  $Revision: 1.6 $
'  Purpose:  Is used for the user to enter in new case details for submission
' ----------------------------------------------------------------------------------
%>

<%
Option Explicit
%>

<!--METADATA TYPE="TypeLib" NAME="Microsoft ActiveX Data Objects 2.5 Library" UUID="{00000205-0000-0010-8000-00AA006D2EA4}" VERSION="2.5"-->
<!--METADATA TYPE="TypeLib" NAME="Microsoft Scripting Runtime" UUID="{420B2830-E718-11CF-893D-00A0C9054228}" VERSION="1.0"-->
<!--METADATA TYPE="TypeLib" NAME="Microsoft CDO for Windows 2000 Library" UUID="{CD000000-8B95-11D1-82DB-00C04FB1625D}" VERSION="1.0"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<!-- #Include File = "Include/Public.asp" -->
<!-- #Include File = "Include/Settings.asp" -->

<!-- #Include File = "Classes/clsAssignment.asp" -->
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


<%


	Dim cnnDB
	Dim I, intUserID, intLastCaseType
	Dim binUserPermMask
	Dim strSQL, strHTML
	Dim rstR
	Dim objContact, objParam, objCollection
	Dim blnIsAdmin, blnIsTech, blnIsUser
	

	' Setup connection to database and identify the user and thier respective permissions

	Set cnnDB = CreateConnection

	intUserID = GetUserID
	binUserPermMask = GetUserPermMask
	

	' Check user permissions for View and Create/Read/Write access

	If (PERM_CREATE_ALL = (PERM_CREATE_ALL And binUserPermMask)) Or (PERM_CREATE_OWN = (PERM_CREATE_OWN And binUserPermMask)) Then

		' Allow case to be created

	Else
	
		' No rights to create new Case
		DisplayError 2, "You do not have the permission to submit a new Case.  Please contact your System Administrator."
	
	End If


	' Determine the page view settings, i.e. what fields will be displayed

	' Initialise view swicth variables, which determine what a user can and can't
	' see.

	blnIsUser = False
	blnIsTech = False
	blnIsAdmin = False

	If VIEW_ADMIN = (VIEW_ADMIN And binUserPermMask) Then
	
		' Admin View access granted
		blnIsAdmin = True
		
	Else

		' Admin View access denied

		If VIEW_TECH = (VIEW_TECH And binUserPermMask) Then
		
			' Tech View access granted
			blnIsTech = True
			
		Else
		
			' Tech View access denied
	
			If VIEW_USER = (VIEW_USER And binUserPermMask) Then

				' User View access granted
				blnIsUser = True

			Else

				' User View access denied

			End If
			
		End If

	End If


	' ------------------------------------------------------------------------
	' XML data island for CaseTypes and associated Categories
	
	strHTML = ""

	strSQL = "SELECT tblCaseTypes.*, tblCategories.* FROM tblCaseTypes " &_
			     "INNER JOIN tblCategories ON tblCaseTypes.CaseTypePK = tblCategories.CaseTypeFK " &_
			     "WHERE tblCaseTypes.IsActive=" & lhd_True & " AND tblCategories.IsActive=" & lhd_True & " " &_
			     "ORDER BY tblCategories.CaseTypeFK, tblCategories.CatPK ASC"
	
	Set rstR = Server.CreateObject("ADODB.Recordset")
	rstR.CursorLocation=adUseClient
	rstR.Open strSQL, cnnDB, adOpenStatic, adLockReadOnly, adCmdText
	
	strHTML = strHTML & "<XML id=""CASETYPES_CATEGORIES"">" & Chr(13)
	strHTML = strHTML & "  <CASETYPES_CATEGORIES>" & Chr(13)

	If rstR.BOF And rstR.EOF Then
	
		' No records returned
		
	Else
	
		While Not rstR.EOF
	
			If rstR.Fields("CaseTypePK") <> intLastCaseType Then
			
				If Not IsEmpty(intLastCaseType) Then
					strHTML = strHTML & "	    </CASETYPE>" & Chr(13)
				Else
					' Do nothing
				End If

				strHTML = strHTML & "    <CASETYPE CaseTypeID=" & Chr(34) & rstR.Fields("CaseTypePK") & Chr(34) & ">" & Chr(13)
				intLastCaseType = rstR.Fields("CaseTypePK")
				
			Else
			
				' Do nothing
			
			End If
  
			strHTML = strHTML & "      <CATEGORY>" & Chr(13)
			strHTML = strHTML & "			   <CatID>" & rstR.Fields("CatPK") & "</CatID>" & Chr(13)
			strHTML = strHTML & "			   <CatName>" & rstR.Fields("CatName") & "</CatName>" & Chr(13)
			strHTML = strHTML & "      </CATEGORY>" & Chr(13)
			
			rstR.MoveNext
  
		WEnd
	
	End If
	
	rstR.Close
	Set rstR = Nothing

	If Not IsEmpty(intLastCaseType)Then
		strHTML = strHTML & "    </CASETYPE>" & Chr(13)
	End If

	strHTML = strHTML & "  </CASETYPES_CATEGORIES>" & Chr(13)
	strHTML = strHTML & "</XML>" & Chr(13)
	
	Response.Write strHTML
	


	' XML data island for CaseTypes and associated Reps

	strHTML = ""

	strSQL = "SELECT tblCaseTypes.CaseTypePK, tblCaseTypes.RepGroupFK, tblGroups.GroupPK, tblGroupMembers.GroupFK, tblGroupMembers.ContactFK, tblContacts.ContactPK, tblContacts.UserName FROM tblContacts " &_
           "INNER JOIN (tblCaseTypes INNER JOIN (tblGroups INNER JOIN tblGroupMembers ON tblGroups.GroupPK = tblGroupMembers.GroupFK) ON tblCaseTypes.RepGroupFK = tblGroups.GroupPK) ON tblContacts.ContactPK = tblGroupMembers.ContactFK " &_
           "WHERE tblCaseTypes.IsActive=" & lhd_True & " AND tblContacts.IsActive=" & lhd_True & " AND tblGroups.IsActive=" & lhd_True & " " &_
           "ORDER BY tblCaseTypes.CaseTypePK, tblCaseTypes.RepGroupFK ASC"

	Set rstR = Server.CreateObject("ADODB.Recordset")
	rstR.CursorLocation=adUseClient
	rstR.Open strSQL, cnnDB, adOpenStatic, adLockReadOnly, adCmdText
	
	strHTML = strHTML & "<XML id=""CASETYPES_REPS"">" & Chr(13)
	strHTML = strHTML & "  <CASETYPES_REPS>" & Chr(13)

	If rstR.BOF And rstR.EOF Then

		' No records

	Else
	
		intLastCaseType = Empty

		While Not rstR.EOF
	
			If rstR.Fields("CaseTypePK") <> intLastCaseType Then
			
				If Not IsEmpty(intLastCaseType) Then
					strHTML = strHTML & "    </CASETYPE>" & Chr(13)
				Else
					' Do nothing
				End If
				
				strHTML = strHTML & "    <CASETYPE CaseTypeID=" & Chr(34) & rstR.Fields("CaseTypePK") & Chr(34) & ">" & Chr(13)
				intLastCaseType = rstR.Fields("CaseTypePK")
						
			Else
			
				' Do nothing
				
			End If
   
			strHTML = strHTML & "      <REP>" & Chr(13)
			strHTML = strHTML & "        <RepID>" & rstR.Fields("ContactPK") & "</RepID>" & Chr(13)
			strHTML = strHTML & "        <RepName>" & rstR.Fields("UserName") & "</RepName>" & Chr(13)
			strHTML = strHTML & "      </REP>" & Chr(13)
			
			rstR.MoveNext
   
		WEnd
		
	End If
	
	rstR.Close
	Set rstR = Nothing

	If Not IsEmpty(intLastCaseType)Then
		strHTML = strHTML & "    </CASETYPE>" & Chr(13)
	Else
		' Do nothing
	End If

	strHTML = strHTML & "  </CASETYPES_REPS>" & Chr(13)
	strHTML = strHTML & "</XML>"
	
	Response.Write strHTML

%>

<HTML>

<SCRIPT language="VBScript">

	Sub LoadAttachmentForm()
	
		window.open "FileAttachment.asp?ID=" & document.frmNew.tbxCaseID.value, "Attachments", "fullscreen=no,toolbar=no,status=no,menubar=no,scrollbars=yes,resizable=no,directories=no,location=no,left=25,top=50,width=550,height=500"

	End Sub


	Sub ListCategoriesAndReps()
	
		Dim XML
		Dim xmlNode, xmlNodes
		Dim objCategoryList, objRepList
	

		' Generate list of associated Categories
	
		Set XML = Document.All("CASETYPES_CATEGORIES")
		Set xmlNodes = XML.SelectNodes("CASETYPES_CATEGORIES/CASETYPE[@CaseTypeID='" & Document.All.cbxCaseType.Value & "']/CATEGORY")

		Set objCategoryList = Document.All("cbxCategory")

		objCategoryList.Options.Length = 1
		objCategoryList.Options(0).Value = 0
		objCategoryList.Options(0).InnerText = "(None)"

		I = 1

		For Each xmlNode In xmlNodes 
			objCategoryList.options.length = objCategoryList.options.length + 1 
			objCategoryList.options(I).Value = xmlNode.SelectSingleNode("CatID").text
			objCategoryList.options(I).InnerText = xmlNode.SelectSingleNode("CatName").text
			
			I = I + 1
		Next

		Set objCategoryList = Nothing


		' Generate list of associated Reps
	
		Set XML = Document.All("CASETYPES_REPS")
		Set xmlNodes = XML.SelectNodes("CASETYPES_REPS/CASETYPE[@CaseTypeID='" & Document.All.cbxCaseType.Value & "']/REP")

		Set objRepList = Document.All("cbxRep")

		objRepList.Options.Length = 1
		objRepList.Options(0).Value = 0
		objRepList.Options(0).InnerText = "(Auto)"

		I = 1

		For Each xmlNode In xmlNodes 
			objRepList.options.length = objRepList.options.length + 1 
			objRepList.options(I).Value = xmlNode.SelectSingleNode("RepID").text
			objRepList.options(I).InnerText = xmlNode.SelectSingleNode("RepName").text
			
			I = I + 1
		Next

		Set objRepList = Nothing
		Set xmlNodes = Nothing
		Set XML = Nothing
	
	End Sub

</SCRIPT>

<%

	Set objContact = New clsContact
	
	objContact.ID = intUserID
	
  If Not objContact.Load Then
	  ' Raise Error, Contact could not be found
	Else
		' Do nothing
  End If

%>

<HEAD>

<META content="MstrHTML 6.00.2600.0" name=GENERATOR></HEAD>

<LINK rel="stylesheet" type="text/css" href="Default.css">

<BODY>
<P align="Center">

<TABLE class=Normal align=center style="WIDTH: 680px" cellSpacing=1 cellPadding=1 width="680" border=0>
	<FORM action="caseNewPost.asp" method="POST" id=frmNew name=frmNew>
	<INPUT type="hidden" name=tbxCaseID id=tbxCaseID value="<%=intUserID & "x" & Right("20" & Year(Now()), 4) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2) & Right("0" & Hour(Now()), 2) & Right("0" & Minute(Now()), 2) & Right("0" & Second(Now()), 2)%>">
	<TR>
		<TD>
		  <%
		  Response.Write DisplayHeader
		  %>
		</TD>
	</TR>
	<TR>
		<TD>
			<TABLE class=Normal cellSpacing=0 cellPadding=1 width="100%" border=0 bgColor=white>
  
				<TR class="lhd_Heading1">
					<TD colspan="5" align="Center"><%=Lang("New_Case")%></TD>
				</TR>
				<TR class="lhd_Heading2">
					<TD colspan="5"><%=Lang("Contact_Detail")%></TD>
				</TR>
				<TR>
					<TD width="20%" align=left class="lhd_Required"><%=Lang("Name")%>:</TD>
					<TD width="27%"><%=objContact.UserName%>&nbsp;&nbsp;<%="( " & objContact.FName & " " & objContact.LName & " )"%></TD>
					<TD width="6%"><INPUT type="hidden" id=tbxContact style="WIDTH: 100%; HEIGHT: 22px" size=4 name=tbxContact value="<%=objContact.ID%>"></TD>
					<TD width="20%"><%=Lang("Phone")%>:</TD>
					<TD width="27%"><%=objContact.OfficePhone%></TD>
				</TR>
				<TR>
					<TD><%=Lang("Department")%>:</TD>
					<TD><INPUT type=hidden name=tbxDeptID value="<%=objContact.DeptID%>"><%=objContact.Dept.DeptName%></TD>
					<TD></TD>
					<TD><%=Lang("Location")%>:</TD>
					<TD><%=objContact.OfficeLocation%></TD>
				</TR>
				<TR>
					<TD align=left><%=Lang("Alternate_EMail")%>:</TD>
					<TD><INPUT id=tbxAltEmail style="WIDTH: 100%; HEIGHT: 22px" size=19 name=tbxAltEmail></TD>
					<TD></TD>
					<TD></TD>
					<TD></TD>
				</TR>
				<TR>
					<TD><%=Lang("Cc")%>:</TD>
					<TD><INPUT id=tbxCc style="WIDTH: 100%; HEIGHT: 22px" size=19 name=tbxCc></TD>
					<TD></TD>
					<TD></TD>
					<TD></TD>
				</TR>
				<%
    		If PERM_CREATE_ALL = (PERM_CREATE_ALL And binUserPermMask) Then
    		%>
				<TR>
					<TD></TD>
					<TD></TD>
					<TD></TD>
					<TD></TD>
					<TD></TD>
				</TR>
				<TR>
					<TD align=left><%=Lang("Or")%> ... <%=Lang("On_Behalf_Of")%>:</TD>
					<TD>
						<SELECT id=cbxContact style="WIDTH: 100%" name=cbxContact>
						  <OPTION value="0" selected>(None)</OPTION>
						  <%
						  Set objCollection = New clsCollection
					
						  objCollection.CollectionType = objCollection.clContact
						  objCollection.Query = "SELECT ContactPK, UserName FROM tblContacts WHERE IsActive=" & lhd_True & " ORDER BY UserName ASC"
					
						  If Not objCollection.Load Then
						
						      Response.Write objCollection.LastError
						      
						  Else
						  	
						  	If objCollection.BOF And objCollection.EOF Then
						  	
						  		' No records
						  	
						  	Else
						
						  		Do While Not objCollection.EOF
						  		%>
						  			<OPTION VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.UserName%></OPTION>
						  		<%
						  			objCollection.MoveNext
						  		Loop
						  		
						  	End If
						  	
						  End If
					
						  Set objCollection = Nothing
						  %>
						</SELECT>
					</TD>
					<TD></TD>
					<TD></TD>
					<TD></TD>
				</TR>
				<%
				Else
				
				  ' Do nothing
				
				End If
				%>
				<TR>
					<TD></TD>
					<TD></TD>
					<TD></TD>
					<TD></TD>
					<TD></TD>
				</TR>
				<TR class="lhd_Heading2">
					<TD colspan="5"><%=Lang("Case_Detail")%></TD>
				</TR>
				<TR>
					<TD class="lhd_Required"><%=Lang("Case_Type")%>:</TD>
					<TD>
						<SELECT id=cbxCaseType style="WIDTH: 100%" name=cbxCaseType onChange="VBScript:ListCategoriesAndReps()">
					    <OPTION VALUE="0" SELECTED>(None)</OPTION>
					    <%
					    Set objCollection = New clsCollection
					    
					    objCollection.CollectionType = objCollection.clCaseType
					    objCollection.Query = "SELECT * FROM tblCaseTypes WHERE IsActive=" & lhd_True & " ORDER BY CaseTypePK ASC"
					    
					    If Not objCollection.Load Then
					    
							' Didn't load
							
						Else
					    
						    If objCollection.BOF And objCollection.EOF Then
					    
								' No records returned
								
							Else
							
								Do While Not objCollection.EOF
								%>
									<OPTION VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.CaseTypeName%></OPTION>
								<%
									objCollection.MoveNext
								Loop
							
							End If
					    
					    End If
					    %>
						</SELECT>
					</TD>
					<TD></TD>
				    <TD class="lhd_Required"><%=Lang("Category")%>:</TD>
					<TD>
						<SELECT id=cbxCategory name=cbxCategory style="WIDTH: 100%" >
						<OPTION  value="0" selected>(None)</OPTION></SELECT>
					</TD>
				</TR>
				<TR>
					<TD><%=Lang("Priority")%>:</TD>
					<TD>
						<SELECT id=cbxPriority style="WIDTH: 100%" name=cbxPriority>
						<%
						Response.Write BuildList("PRIORITY_LIST", Application("DEFAULT_PRIORITY"))
				        %>
						</SELECT>
					</TD>
					<TD></TD>
					<TD></TD>
					<TD></TD>
				</TR>
				<%
				If blnIsAdmin = True OR blnIsTech = True Then
				%>
				<TR>
					<TD><%=Lang("Status")%>:</TD>
					<TD>
						<SELECT id=cbxStatus name=cbxStatus style="WIDTH: 100%">
						<%
						Response.Write BuildList("STATUS_LIST", Application("DEFAULT_STATUS"))
				        %>
						</SELECT>
					</TD>
					<TD></TD>
					<TD></TD>
					<TD></TD>
				</TR>
				<TR>
					<TD><%=Lang("Assignment")%>:</TD>
					<TD>
						<SELECT id=cbxRep name=cbxRep style="WIDTH: 100%">
						<OPTION value="0" selected>(Auto)</OPTION>
						</SELECT>
					</TD>
					<TD><INPUT type="hidden" id=tbxEnteredByID style="WIDTH: 100%; HEIGHT: 22px" size=4 name=tbxEnteredByID value="<%=objContact.ID%>"></TD>
					<TD><%=Lang("Entered_By")%>:</TD>
					<TD><%=objContact.UserName%></TD>
				</TR>
				<%
				Else
					' blnIsUser View
				%>				
					<SELECT id=cbxStatus name=cbxStatus style="DISPLAY: NONE; WIDTH: 100%">
						<%
						Response.Write BuildList("STATUS_LIST", Application("DEFAULT_STATUS"))
						%>
					</SELECT>
					<SELECT id=cbxRep name=cbxRep style="DISPLAY: NONE; WIDTH: 100%">
						<OPTION value="0" selected>(Auto)</OPTION>
					</SELECT>
					<INPUT type="hidden" id=tbxEnteredByID style="DISPLAY: NONE; WIDTH: 100%; HEIGHT: 22px" size=4 name=tbxEnteredByID value="<%=objContact.ID%>">
				<%
				End If
				%>
				<TR>
					<TD class="lhd_Required"><%=Lang("Title")%>:</TD>
				  <%
				  If Application("ENABLE_ATTACHMENTS") = 1 Then
				  %>
					<TD colspan=3><INPUT id=tbxTitle style="WIDTH: 100%" name=tbxTitle></TD>
					<TD valign=top align=right>
  					<INPUT style="WIDTH: 150px; BACKGROUND-COLOR: white" type="button" value="Attachments" name=btnAttachments id=btnAttachments onClick="VBScript:LoadAttachmentForm()">
					</TD>
  				<%
  				Else
  				%>
					<TD colspan=4><INPUT id=tbxTitle style="WIDTH: 100%" name=tbxTitle></TD>
					<%
  				End If
  				%>
				</TR>
				<TR>
					<TD vAlign=top><%=Lang("Detailed_Description")%>:</TD>
					<TD colspan=4><TEXTAREA id=txtDescription style="SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-TRACK-COLOR: white; WIDTH: 100%; HEIGHT: 140px" name=txtDescription></TEXTAREA></TD>
				</TR>
				<%
				If blnIsAdmin=True OR blnIsTech=True Then
				%>
				<TR>
					<TD vAlign=top><%=Lang("Notes")%>:</TD>
					<TD colspan=4><TEXTAREA id=txtNotes style="SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-TRACK-COLOR: white; WIDTH: 100%; HEIGHT: 70px" name=txtNotes></TEXTAREA></TD>
				</TR>
				<TR>
					<TD></TD>
					<TD colspan=2><INPUT id=chkPrivateNote type=checkbox name=chkPrivateNote style="LEFT: 4px; TOP: 3px">&nbsp;<%=Lang("Private_Note")%></TD>
					<TD><%=Lang("Time_Spent")%>:</TD>
					<TD><INPUT id=tbxMinutesSpent style="WIDTH: 100%" name=tbxMinutesSpent></TD>
				</TR>
				<TR>
					<TD></TD>
					<TD></TD>
					<TD></TD>
					<TD></TD>
					<TD></TD>
				</TR>
				<TR class="lhd_Heading2">
					<TD colspan="5"><%=Lang("Resolution")%></TD>
				</TR>
				<TR>
					<TD vAlign=top><%=Lang("Resolution")%>:</TD>
					<TD colspan=4><TEXTAREA id=txtResolution style="SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-TRACK-COLOR: white; WIDTH: 100%; HEIGHT: 140px" name=txtResolution></TEXTAREA></TD>
				</TR>
				<TR>
					<TD></TD>
					<TD colspan=2><INPUT id=chkNotifyUser type=checkbox name=chkNotifyUser>&nbsp;<%=Lang("Notify_User")%></TD>
					<TD></TD>
					<TD></TD>
				</TR>
				<%
				Else
				
					' blnIsUser View
				%>

					<TEXTAREA id=txtNotes style="DISPLAY: NONE; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-TRACK-COLOR: white; WIDTH: 100%; HEIGHT: 70px" name=txtNotes></TEXTAREA>
					<TEXTAREA id=txtResolution style="DISPLAY: NONE; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-TRACK-COLOR: white; WIDTH: 100%; HEIGHT: 140px" name=txtResolution></TEXTAREA>

					<INPUT id=chkPrivateNote style="DISPLAY: NONE; LEFT: 4px; TOP: 3px" type=checkbox name=chkPrivateNote >
					<INPUT id=tbxMinutesSpent style="DISPLAY: NONE; WIDTH: 100%" name=tbxMinutesSpent>
					<INPUT id=chkNotifyUser style="DISPLAY: NONE;" type=checkbox name=chkNotifyUser>
				
				<%
				End If
				%>
				<TR>
					<TD></TD>
					<TD></TD>
					<TD></TD>
					<TD></TD>
					<TD></TD>
				</TR>
				<TR>
					<TD></TD>
					<TD colspan=4 align=right><INPUT style="WIDTH: 80px; BACKGROUND-COLOR: white" id=btnSubmit type=submit align=right value="<%=Lang("Submit")%>" name=btnSubmit></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD>
		  <%
		  Response.Write DisplayFooter
		  %>
		</TD>
	</TR>
</FORM>
</TABLE>
</P>
</BODY>

</HTML>

<%
Set objContact = Nothing

cnnDB.Close
Set cnnDB = Nothing
%>
