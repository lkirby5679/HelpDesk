<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: caseModify.asp
'  Date:     $Date: 2004/03/29 04:27:48 $
'  Version:  $Revision: 1.10 $
'  Purpose:  Used to allow case details to be modified
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
<!-- #Include File = "Classes/clsEMailMsg.asp" -->
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
	Dim I
	Dim intUserID, intLastCaseType
	Dim lngCaseID
	Dim binUserPermMask, binRequiredPerm, blnPrintView
	Dim objNote, objCase, objContact, objCollection
	Dim strSQL, strHTML
	Dim rstR, rstReps, rstAttachments
	Dim blnADMIN, blnTECH, blnUSER, blnCanRead, blnCanModify, blnMatch

	' Setup connection to database and get user details
	Set cnnDB = CreateConnection

	intUserID = GetUserID
	binUserPermMask = GetUserPermMask
	
	' Determine the page view settings, i.e. what fields will be displayed
	'
	' Initialise view switch variables, which determine what a user can and can't
	' see.

	blnADMIN = False
	blnTECH = False
	blnUSER = False


	If VIEW_ADMIN = (VIEW_ADMIN And binUserPermMask) Then
	
		' Admin View access granted
		
		blnADMIN = True
		
	Else

		' Admin View access denied

		If VIEW_TECH = (VIEW_TECH And binUserPermMask) Then
		
			' Tech View access granted

			blnTECH = True
			
		Else
		
			' Tech View access denied
	
			If VIEW_USER = (VIEW_USER And binUserPermMask) Then

				' User View access granted
				blnUSER = True

			Else

				' User View access denied

			End If
			
		End If

	End If


	If Request.QueryString.Count > 0 Then
		lngCaseID = CInt(Request.QueryString("ID"))
	Else
		lngCaseID = CInt(Request.Form("tbxID"))
	End If
	
	If Request.QueryString("PrintView") = 1 Then
		blnPrintView = True
	Else
		blnPrintView = False
	End If

	
	Set objCase = New clsCase
	
	objCase.ID = lngCaseID
	
	If Not objCase.Load Then
	
		' Case no found or loaded
		
		DisplayError 3, "Case Not Found"
		
	Else
	
		' Case loaded. Now check whether or not the logged in user is allowed to
		' view/modify this case.
		
		'-----------------------------------------------------------------------
		
		blnCanRead = False
		
		If PERM_READ_ALL = (PERM_READ_ALL And binUserPermMask) Then

			' GRANTED: user has access to view all cases
			blnCanRead = True			

		Else
			
			' DENIED: user only has access to their own cases
			
			' Check if this Case belongs to this user.
			If (objCase.ContactID = intUserID) Or (InStr(1, objCase.Cc, Session("Username")) > 0) Then

				If PERM_READ_OWN = (PERM_READ_OWN And binUserPermMask) Then
				
					' GRANTED: user has access to view this case
					blnCanRead = True
				
				Else
				
					' DENIED: user has no read access to view this case
				
				End If

			Else

				If objCase.RepID = intUserID Then
				
					If PERM_READ_ASSIGNED = (PERM_READ_ASSIGNED And binUserPermMask) Then
					
						' GRANTED: This user is the assiged rep to this case
						blnCanRead = True
						
					Else
	
						' DENIED: The assigned Rep has no read access this case
					
					End If
				
				Else
				
					' Finally we need to check if the person logged in has PERM_ACCESS_TECH
					' that they belong to the assigned RepGroupPK.
					If objCase.Group.IsMember(intUserID) Then

						If PERM_READ_GROUP = (PERM_READ_GROUP And binUserPermMask) Then
					
							' GRANTED: This user is the assiged rep to this case
							blnCanRead = True
							
						Else
	
							' DENIED: The assigned Rep has no read access this case
					
						End If
					
					Else
					
						' DENIED: user doesn't have access to view this case
								
					End If
					
				End If

			End If
			
		End If
	
		'-----------------------------------------------------------------------
		
		blnCanModify = CanModifyCase(objCase, intUserID, binUserPermMask)
		
		'-----------------------------------------------------------------------

	End If
	
	Set objCase = Nothing


	' Now check  if the user logged in can read to case
	
	If blnCanRead = True Then
	
		' Allow user to read case
		
	Else
	
		DisplayError 4, ""
	
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
	strHTML = strHTML & "<CASETYPES_CATEGORIES>" & Chr(13)

	If rstR.BOF And rstR.EOF Then
	
		' No records returned
		
	Else
	
		While Not rstR.EOF
	
			If rstR.Fields("CaseTypePK") <> intLastCaseType Then
			
				If Not IsEmpty(intLastCaseType) Then
					strHTML = strHTML & "	</CASETYPE>" & Chr(13)
				Else
					' Do nothing
				End If

				strHTML = strHTML & "	<CASETYPE CaseTypeID=" & Chr(34) & rstR.Fields("CaseTypePK") & Chr(34) & ">" & Chr(13)
				intLastCaseType = rstR.Fields("CaseTypePK")
				
			Else
			
				' Do nothing
			
			End If
  
			strHTML = strHTML & "		<CATEGORY>" & Chr(13)
			strHTML = strHTML & "			<CatID>" & rstR.Fields("CatPK") & "</CatID>" & Chr(13)
			strHTML = strHTML & "			<CatName>" & rstR.Fields("CatName") & "</CatName>" & Chr(13)
			strHTML = strHTML & "		</CATEGORY>" & Chr(13)
			
			rstR.MoveNext
  
		WEnd
	
	End If
	
	rstR.Close
	Set rstR = Nothing

	If Not IsEmpty(intLastCaseType)Then
		strHTML = strHTML & "	</CASETYPE>" & Chr(13)
	End If

	strHTML = strHTML & "</CASETYPES_CATEGORIES>" & Chr(13)
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
	strHTML = strHTML & "<CASETYPES_REPS>" & Chr(13)

	If rstR.BOF And rstR.EOF Then

		' No records

	Else
	
		intLastCaseType = Empty

		While Not rstR.EOF
	
			If rstR.Fields("CaseTypePK") <> intLastCaseType Then
			
				If Not IsEmpty(intLastCaseType) Then
					strHTML = strHTML & "	</CASETYPE>" & Chr(13)
				Else
					' Do nothing
				End If
				
				strHTML = strHTML & "	<CASETYPE CaseTypeID=" & Chr(34) & rstR.Fields("CaseTypePK") & Chr(34) & ">" & Chr(13)
				intLastCaseType = rstR.Fields("CaseTypePK")
						
			Else
			
				' Do nothing
				
			End If
   
			strHTML = strHTML & "		<REP>" & Chr(13)
			strHTML = strHTML & "			<RepID>" & rstR.Fields("ContactPK") & "</RepID>" & Chr(13)
			strHTML = strHTML & "			<RepName>" & rstR.Fields("UserName") & "</RepName>" & Chr(13)
			strHTML = strHTML & "		</REP>" & Chr(13)
			
			rstR.MoveNext
   
		WEnd
		
	End If
	
	rstR.Close
	Set rstR = Nothing

	If Not IsEmpty(intLastCaseType)Then
		strHTML = strHTML & "	</CASETYPE>" & Chr(13)
	Else
		' Do nothing
	End If

	strHTML = strHTML & "</CASETYPES_REPS>" & Chr(13)
	strHTML = strHTML & "</XML>"
	
	Response.Write strHTML

%>

<HTML>

<SCRIPT language="VBScript">

	Sub LoadAttachmentForm()
	
		window.open "fileattachment.asp?id=" & document.frmModify.tbxCaseID.value, "Attachments", "fullscreen=no,toolbar=yes,status=no,menubar=no,scrollbars=yes,resizable=no,directories=no,location=no,left=25,top=50,width=550,height=500"
	
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

	Set objCase = New clsCase

	objCase.ID = lngCaseID

  If Not objCase.Load Then
		' Raise Error, as no record found
	Else
		' Do nothing
  End If
    
%>

<HEAD>

<META content="MstrHTML 6.00.2600.0" name=GENERATOR></HEAD>

<LINK rel="stylesheet" type="text/css" href="default.css">

<BODY>
<P align=center>
<TABLE class=Normal align=center style="WIDTH: 680px" cellSpacing=1 cellPadding=1 width="680" border=0>
<FORM action="caseModifyPost.asp" method="POST" id=frmModify name=frmModify>
  <INPUT type="hidden" name=tbxUpdate id=tbxUpdate value="1">
  <INPUT type="hidden" name=tbxCaseID id=tbxCaseID value=<%=objCase.ID%>>

	<TR>
		<TD>
		<%
		Response.Write DisplayHeader
		%>
		</TD>
	</TR>
  <TR>
    <TD>

      <%
      If objCase.StatusID <> Application("STATUS_CLOSED") And objCase.StatusID <> Application("STATUS_CANCELLED") And blnCanModify = True And blnPrintView = False Then
      %>

        <INPUT type="hidden" name="tbxReOpen" id="tbxReOpen" value="0">
        <TABLE class=Normal cellSpacing=0 cellPadding=1 width="100%" border=0 style="WIDTH: 100%" bgColor=white>
	        <TR class="lhd_Heading1">
	        	<TD colspan=5 align=center><%=Lang("Modify")%>&nbsp<%=Lang("Case")%>&nbsp;&nbsp;#<%=objCase.ID%></TD>
	        </TR>
	        <TR class="lhd_Heading2">
	        	<TD colspan=5><%=Lang("Contact_Detail")%></TD>
	        </TR>
	        <TR>
	        	<TD width="20%" align=left class="lhd_Required"><%=Lang("Name")%>:</TD>
	        	<TD width="27%"><%=objCase.Contact.UserName%></TD>
	        	<TD width="6%" ><INPUT type="hidden" id=tbxContactID style="WIDTH: 100%; HEIGHT: 22px" size="4" name="tbxContactID" value="<%=objCase.Contact.ID%>"></TD>
	        	<TD width="20%"><%=Lang("Phone")%>:</TD>
	        	<TD width="27%"><%=objCase.Contact.OfficePhone%></TD>
	        </TR>
	        <TR>
	        	<TD><%=Lang("Department")%>:</TD>
	        	<TD><%=objCase.Dept.DeptName%></TD>
        		<TD></TD>
	        	<TD><%=Lang("Location")%>:</TD>
	        	<TD><%=objCase.Contact.OfficeLocation%></TD>
	        </TR>
	        <TR>
	        	<TD align=left><%=Lang("Alternate_EMail")%>:</TD>
	        	<TD><INPUT id=tbxAltEmail style="WIDTH: 100%; HEIGHT: 22px" size=19 name=tbxAltEmail value="<%=objCase.AltEMail%>"></TD>
	        	<TD></TD>
	        	<TD></TD>
	        	<TD></TD>
	        </TR>
	        <TR>
	        	<TD><%=Lang("Cc")%>:</TD>
	        	<TD><INPUT id=tbxCc style="WIDTH: 100%; HEIGHT: 22px" size=19 name=tbxCc value="<%=objCase.Cc%>"></TD>
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
	        <TR class="lhd_Heading2">
	        	<TD colspan=5><%=Lang("Case_Detail")%></TD>
	        </TR>
	        <%
	        If blnADMIN = True OR blnTECH = True Then
	        %>
	        <TR>
	        	<TD class="lhd_Required"><%=Lang("Case_Type")%>:</TD>
	        	<TD>
	        		<SELECT id=cbxCaseType style="WIDTH: 100%" name=cbxCaseType onChange="VBScript:ListCategoriesAndReps()">
	        	    <OPTION SELECTED VALUE="0">(None)</OPTION>
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

	        			  		If objCollection.Item.ID = objCase.CaseTypeID Then
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
	        	    %>
	        		</SELECT>
	        	<TD></TD>
	        	<TD class="lhd_Required"><%=Lang("Category")%>:</TD>
	        	<TD>
	        		<SELECT id=cbxCategory name=cbxCategory style="WIDTH: 100%">
	        	    <OPTION VALUE="0">(None)</OPTION>
	        	    <%
	        	    If Not IsNull(objCase.CaseTypeID) Then
	        	    
	        			  Set objCollection = New clsCollection
	        			  			    
	        			  objCollection.CollectionType = objCollection.clCategory
	        			  objCollection.Query = "SELECT * FROM tblCategories WHERE IsActive=" & lhd_True & " AND CaseTypeFK = " & objCase.CaseTypeID
	        			  			    
	        			  If Not objCollection.Load Then
	        			  			    
	        			  	' Didn't load
	        			  					
	        			  Else
	        			  			    
	        			      If objCollection.BOF And objCollection.EOF Then
	        			  			    
	        			  		' No records returned
	        			  						
	        			  	Else
	        			  					
	        			  		Do While Not objCollection.EOF

	        			  			If objCollection.Item.ID = objCase.CatID Then
	        			  			%>
	        			  			<OPTION SELECTED VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.CatName%></OPTION>
	        			  			<%
	        			  			Else
	        			  			%>
	        			  			<OPTION VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.CatName%></OPTION>
	        			  			<%
	        			  			End If
	        			  			
	        			  			objCollection.MoveNext
	        			  			
	        			  		Loop
	        			  					
	        			  	End If
	        			  			    
	        			  End If
	        			
	        		  Else
	        		
	        		  	' Do Nothing
	        		  	
	        		  End If
	        	    %>
	        	    </SELECT>
	        	</TD>
	        </TR>
          <TR>
	        <TD><%=Lang("Priority")%>:</TD>
	        <TD>
	        	<SELECT id=cbxPriority style="WIDTH: 100%" name=cbxPriority>
	        	<%
	        	Response.Write BuildList("PRIORITY_LIST", objCase.PriorityID)
	            %>
	        	</SELECT>
	        </TD>
            <TD></TD>
            <TD><%=Lang("Start_Date")%>:</TD>
            <TD><%=DisplayDateTime(objCase.RaisedDate)%></TD>
          <TR>
	        <TD><%=Lang("Status")%>:</TD>
	        <TD>
	        	<SELECT id=cbxStatus name=cbxStatus style="WIDTH: 100%">
	        	<%
	        	Response.Write BuildList("STATUS_LIST", objCase.StatusID)
	            %>
	        	</SELECT>
            <TD></TD>
            <TD><%=Lang("Closed_Date")%>:</TD>
            <TD><%=DisplayDateTime(objCase.ClosedDate)%></TD></TR>
          </TR>
          <TR>
            <TD><%=Lang("Assignment")%>:</TD>
            <TD>
	        	<SELECT id=cbxRep name=cbxRep style="WIDTH: 100%">
	        		<OPTION selected value="0">(Unassigned)</OPTION>
                    <%
                    If Not IsNull(objCase.CaseTypeID) Then
                    
	        			      strSQL = "SELECT tblCaseTypes.CaseTypePK, tblCaseTypes.RepGroupFK, tblGroups.GroupPK, tblGroupMembers.GroupFK, tblGroupMembers.ContactFK, tblContacts.ContactPK, tblContacts.UserName FROM tblContacts " &_
	        			      		     "INNER JOIN (tblCaseTypes INNER JOIN (tblGroups INNER JOIN tblGroupMembers ON tblGroups.GroupPK = tblGroupMembers.GroupFK) ON tblCaseTypes.RepGroupFK = tblGroups.GroupPK) ON tblContacts.ContactPK = tblGroupMembers.ContactFK " &_
            		      				 "WHERE tblCaseTypes.IsActive=" & lhd_True & " AND tblCaseTypes.CaseTypePK=" & objCase.CaseTypeID & " AND tblContacts.IsActive=" & lhd_True & " AND tblGroups.IsActive=" & lhd_True

	        		      	Set rstReps = CreateObject("ADODB.RecordSet")
	        		      	rstReps.Open  strSQL, cnnDB

	        		      	If rstReps.EOF And rstReps.BOF Then

	        		      		' No records

	        		      	Else
                          
	              			    blnMatch = False
	              			    
	        		      	    Do While Not rstReps.EOF
	        		      	    
	        		      			  If rstReps("ContactPK") = objCase.RepID Then
	        		      			    blnMatch = True
	        		      			    %>
	        		      			  	<OPTION SELECTED VALUE="<%=rstReps("ContactPK")%>"><%=rstReps("UserName")%></OPTION>
	        		      			    <%
	        		      			  Else
	        		      			    %>
	        		      			  	<OPTION VALUE="<%=rstReps("ContactPK")%>"><%=rstReps("UserName")%></OPTION>
	        		      			    <%
	        		      			  End If
	        		      			
	        		      			  rstReps.MoveNext
	        		      			
	        		      	    Loop
	        		      	    
	        		      	    If blnMatch = False And objCase.RepID > 0 Then
                  			    %>
                  			    <OPTION selected value="<%=objCase.RepID%>"><%=objCase.Rep.UserName%></OPTION>
                  			    <%
                  			  Else
                  			    ' Do nothing
                  			  End If

	        		      	End If
                          
	        		      	Set rstReps = Nothing
	        		
	        		      Else
	        		
	        		      	' Do nothing
	        		
	        		      End If
                    %>
                </SELECT>
            </TD>
            <TD></TD>
            <TD><%=Lang("Entered_By")%>:</TD>
            <TD><%=objCase.EnteredBy.UserName%></TD>
            </TR>
	        <%
	        Else
	        %>	
	        <TR>
	        	<SELECT style="DISPLAY: NONE; WIDTH: 100%" id=cbxCaseType name=cbxCaseType>
	        		<OPTION selected value="<%=objCase.CaseTypeID%>"><%=objCase.CaseType.CaseTypeName%></OPTION>
	        	</SELECT>
	        	<SELECT style="DISPLAY: NONE; WIDTH: 100%" id=cbxCategory name=cbxCategory>
	        		<OPTION selected value="<%=objCase.CatID%>"><%=objCase.Cat.CatName%></OPTION>
	        	</SELECT>
	        	<TD class="lhd_Required"><%=Lang("Case_Type")%>:</TD>
	        	<TD><%=objCase.CaseType.CaseTypeName%></TD>
	        	<TD></TD>
	        	<TD class="lhd_Required"><%=Lang("Category")%>:</TD>
	        	<TD><%=objCase.Cat.CatName%></TD>
	        </TR>
	        <TR>
	        	<SELECT style="DISPLAY: NONE; WIDTH: 100%" id=cbxPriority name=cbxPriority>
	        		<OPTION selected value="<%=objCase.PriorityID%>"><%=objCase.Priority.ItemName%></OPTION>
	        	</SELECT>
	        	<TD><%=Lang("Priority")%>:</TD>
	        	<TD><%=objCase.Priority.ItemName%></TD>
	        	<TD></TD>
	        	<TD></TD>
	        	<TD></TD>
	        </TR>
	        <TR>
	        	<SELECT style="DISPLAY: NONE; WIDTH: 100%" id=cbxStatus name=cbxStatus>
	        		<OPTION selected value="<%=objCase.StatusID%>"><%=objCase.Status.ItemName%></OPTION>
	        	</SELECT>
	        	<TD><%=Lang("Status")%>:</TD>
	        	<TD><%=objCase.Status.ItemName%></TD>
	        	<TD></TD>
	        	<TD></TD>
	        	<TD></TD>
	        </TR>
	        <TR>
	        	<SELECT style="DISPLAY: NONE; WIDTH: 100%" id=cbxRep name=cbxRep>
	        		<%
	        		If Len(objCase.RepID) > 0 Then
	        		%>
	        			<OPTION selected value="<%=objCase.RepID%>"><%=objCase.Rep.UserName%></OPTION>
	        		<%
	        		Else
	        		%>
	        			<OPTION selected value="0">(Unassigned)</OPTION>
	        		<%
	        		End If
	        		%>
	        	</SELECT>
	        	<TD><%=Lang("Assignment")%>:</TD>
	        	<TD>
	        		<%
	        		If IsEmpty(objCase.RepID) Or IsNull(objCase.RepID) Or objCase.RepID = 0 Then
	        			Response.Write "Unassigned"
	        		Else
	        			Response.Write objCase.Rep.UserName
	        		End If
	        		%>
	        	</TD>
	        	<TD></TD>
	        	<TD><%=Lang("Entered_By")%>:</TD>
	        	<TD><%=objCase.EnteredBy.UserName%></TD>
	        </TR>
	        <%	
	        End If
	        %>
          <TR>
            <TD class="lhd_Required"><%=Lang("Title")%>:</TD>
	        	<%
				    If Application("ENABLE_ATTACHMENTS") = 1 Then
				    %>
            <TD colspan=3><INPUT id="Text1" style="WIDTH: 100%" name=tbxTitle value="<%=objCase.Title%>"></TD>
	        	<TD valign=top align=right>
              <INPUT style="WIDTH: 150px; BACKGROUND-COLOR: white" type="button" value="Attachments" name=btnAttachments id=btnAttachments onClick="VBScript:LoadAttachmentForm()">
            </TD>
  				  <%
  				  Else
  				  %>
            <TD colspan=4><INPUT id="Text1" style="WIDTH: 100%" name=tbxTitle value="<%=objCase.Title%>"></TD>
            <%
  				  End If
  				  %>
          </TR>
    		  <%

					Set rstAttachments = Server.CreateObject("ADODB.Recordset")
							
					rstAttachments.Open "SELECT * FROM tblFiles WHERE CaseFK='" & lngCaseID & "'", cnnDB
							
					If rstAttachments.BOF And rstAttachments.EOF Then
							
						' No records found
								
					Else
							
					  strHTML = ""
							
						Do 

              If Len(strHTML) > 0 Then
  							strHTML = strHTML & ",&nbsp;<A href=""FileView.asp?ID=" & rstAttachments.Fields("FilePK") & """ target=""_blank"">" &  rstAttachments.Fields("FileName") & "&nbsp;(" & CStr(Int(rstAttachments.Fields("FileSize")/1024)) & "k)" & "</A>"
  					  Else
  							strHTML = strHTML & "<A href=""FileView.asp?ID=" & rstAttachments.Fields("FilePK") & """ target=""_blank"">" &  rstAttachments.Fields("FileName") & "&nbsp;(" & CStr(Int(rstAttachments.Fields("FileSize")/1024)) & "k)" & "</A>"
  					  End If
							rstAttachments.MoveNext

						Loop Until rstAttachments.EOF

            %>								
            <TR>
              <TD></TD>
              <TD colspan="4">
                <FONT style="FONT-SIZE: 8pt">
                  <%
					  			Response.Write strHTML
					        %>
					      </FONT>
              </TD>
            </TR>
            <%							

					End If
							
					rstAttachments.Close
					Set rstAttachments = Nothing
					%>
          <TR>
            <TD vAlign=top><%=Lang("Detailed_Description")%>:</TD>
            <TD colspan=4><TEXTAREA id=txtDescription style="SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-TRACK-COLOR: white; WIDTH: 100%; HEIGHT: 140px" name=txtDescription><%=objCase.Description%></TEXTAREA></TD>
            </TR>
          <TR>
            <TD colspan="5"></TD>
          </TR>
          <TR>
            <TD vAlign=top><%=Lang("Notes")%>:</TD>
            <TD colspan=4>
	        	<%

	        	' Need to load all of the notes for this case
	        	
	        	strHTML = ""
	        	
	        	If objCase.CaseNotes.BOF And objCase.CaseNotes.EOF Then
	        	
	        		' No case notes to list
	        		
	        	Else
	        	
	        		While Not objCase.CaseNotes.EOF
	        		
	        			If objCase.CaseNotes.Item.IsPrivate Then
	        			
	        				If blnTECH = True or blnADMIN Then
	        				
	        					strHTML = strHTML & "<FONT Size=1><B>[" & DisplayDateTime(objCase.CaseNotes.Item.AddDate) & "&nbsp;&nbsp;" & objCase.CaseNotes.Item.Owner.UserName & " - PRIVATE]</B><BR>"
	        					strHTML = strHTML & Replace(objCase.CaseNotes.Item.Note, Chr(13) & Chr(10), "<BR>") & "<BR></FONT>"
	        				
	        				Else
	        				
	        					' Do nothing, as the user isn't allow to see this note
	        				
	        				End If
	        			
	        			Else
	        			
	        				strHTML = strHTML & "<FONT Size=1><B>[" & DisplayDateTime(objCase.CaseNotes.Item.AddDate) & "&nbsp;&nbsp;" & objCase.CaseNotes.Item.Owner.UserName & "]</B><BR>"
	        				strHTML = strHTML & Replace(objCase.CaseNotes.Item.Note, Chr(13) & Chr(10), "<BR>") & "<BR></FONT>"
	        			
	        			End If

	        			objCase.CaseNotes.MoveNext
	        		
	        		WEnd
	        		
	        	End If

	        	Response.Write strHTML
            
	        	%>
	        	<TEXTAREA id=txtNotes style="SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-TRACK-COLOR: white; WIDTH: 100%; HEIGHT: 70px" name=txtNotes></TEXTAREA>
            </TD>
            </TR>
	        <%
	        If blnADMIN = True OR blnTECH = True Then
	        %>
	        <TR>
	        	<TD></TD>
	        	<TD colspan=2><INPUT style="LEFT: 4px; TOP: 3px" id=chkPrivateNote type=checkbox name=chkPrivateNote>&nbsp;<%=Lang("Private_Note")%></TD>
	        	<TD><%=Lang("Time_Spent")%>:</TD>
	        	<TD><INPUT id=tbxMinutesSpent style="WIDTH: 100%" name=tbxMinutesSpent></TD>
	        </TR>
	        <%
	        Else
	        %>
	
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
	        <TR class="lhd_Heading2">
	        	<TD colspan=5><%=Lang("Resolution")%></TD>
	        </TR>
	        <%
	        If blnADMIN = True OR blnTECH = True Then
	        %>
	        	<TR>
	        	  <TD vAlign=top><%=Lang("Resolution")%>:</TD>
	        	  <TD colspan=4><TEXTAREA id=txtResolution style="SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-TRACK-COLOR: white; WIDTH: 100%; HEIGHT: 140px" name=txtResolution><%=objCase.Resolution%></TEXTAREA></TD>
	          </TR>
	        <%
	        Else
	        %>
	        	<TR>
	        	  <TD vAlign=top><%=Lang("Resolution")%>:</TD>
	        	  <TD colspan=4><TEXTAREA id=txtResolution style="SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-TRACK-COLOR: white; WIDTH: 100%; HEIGHT: 140px" name=txtResolution readonly><%=objCase.Resolution%></TEXTAREA></TD>
	          </TR>
	        <%
	        End If
	        %>

	        <%
	        If blnADMIN = True OR blnTECH = True Then
	        %>
	        <TR>
	        	<TD></TD>
	        	<TD colspan=2><INPUT id=chkNotifyUser type=checkbox name=chkNotifyUser>&nbsp;<%=Lang("Notify_User")%></TD>
	        	<TD></TD>
	        	<TD></TD>
	        </TR>
	        <%
	        Else
	        %>
	
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
	        <%
	        'If blnCanModify = True Then
	        %>
	        <TR>
	        	<%
        		If blnPrintView = False Then
        		%>
        		  <td></td>
        		  <TD align=Left><a href="caseModify.asp?ID=<%=lngCaseID%>&PrintView=1"><%=Lang("Printer_Friendly_View")%></a></TD>
        		<%
        		Else
        		%>
        		  <TD colspan=2></TD>  
        		<%
        	  End If
        		%>
	        	<TD colspan=3 align=right><INPUT style="WIDTH: 80px; BACKGROUND-COLOR: white" type="submit" value="<%=Lang("Update")%>" name=btnUpdate id=btnUpdate></TD>
	        </TR>
	        <%
	        'Else
	        	' Do nothing
	        'End If
	        %>
        </TABLE>

      <%
      Else  ' Display read only case
      %>

        <INPUT type="hidden" name="tbxReOpen" id="tbxReOpen" value="1">
        <TABLE class="Normal" cellSpacing="0" cellPadding="1" width="100%" border="0" style="WIDTH: 100%" bgColor="white">
        	<TR class="lhd_Heading1">
        		<TD colspan=5 align=center><%=Lang("Modify")%>&nbsp<%=Lang("Case")%>&nbsp;&nbsp;#<%=objCase.ID%></TD>
        	</TR>
        	<TR class="lhd_Heading2">
        		<TD colspan=5><%=Lang("Contact_Detail")%></TD>
        	</TR>
        	<TR>
        		<TD width="20%" align=left class="lhd_Required"><%=Lang("Name")%>:</TD>
        		<TD width="27%"><%=objCase.Contact.UserName%></TD>
        		<TD width="6%" ></TD>
        		<TD width="20%"><%=Lang("Phone")%>:</TD>
        		<TD width="27%"><%=objCase.Contact.OfficePhone%></TD>
        	</TR>
        	<TR>
        		<TD><%=Lang("Department")%>:</TD>
        		<TD><%=objCase.Dept.DeptName%></TD>
        		<TD></TD>
        		<TD><%=Lang("Location")%>:</TD>
        		<TD><%=objCase.Contact.OfficeLocation%></TD>
        	</TR>
        	<TR>
        		<TD align=left><%=Lang("Alternate_EMail")%>:</TD>
        		<TD><%=objCase.AltEMail%></TD>
        		<TD></TD>
        		<TD></TD>
        		<TD></TD>
        	</TR>
        	<TR>
        		<TD><%=Lang("Cc")%>:</TD>
        		<TD><%=objCase.Cc%></TD>
        		<TD></TD>
        		<TD></TD>
        		<TD></TD>
        	</TR>
        	<%
        	If blnPrintView = False Then
        	%>
        	<TR>
        		<TD colspan="5"></TD>
        	</TR>
        	<%
        	Else
        	  ' Do nothing
        	End If
        	%>
        	<TR class="lhd_Heading2">
        		<TD colspan=5><%=Lang("Case_Detail")%></TD>
        	</TR>
        	<TR>
        		<TD class="lhd_Required"><%=Lang("Case_Type")%>:</TD>
        		<TD><%=objCase.CaseType.CaseTypeName%></TD>
        		<TD></TD>
        		<TD class="lhd_Required"><%=Lang("Category")%>:</TD>
        		<TD><%=objCase.Cat.CatName%></TD>
        	</TR>
        	<TR>
        		<TD><%=Lang("Priority")%>:</TD>
        		<TD><%=objCase.Priority.ItemName%></TD>
        		<TD></TD>
        		<TD></TD>
        		<TD></TD>
        	</TR>
        	<TR>
        		<TD><%=Lang("Status")%>:</TD>
        		<TD><%=objCase.Status.ItemName%></TD>
        		<TD></TD>
        		<TD></TD>
        		<TD></TD>
        	</TR>
        	<TR>
        		<TD><%=Lang("Assignment")%>:</TD>
        		<TD>
        			<%
        			If IsEmpty(objCase.RepID) Or IsNull(objCase.RepID) Or objCase.RepID = 0 Then
        				Response.Write "Unassigned"
        			Else
        				Response.Write objCase.Rep.UserName
        			End If
        			%>
        		</TD>
        		<TD></TD>
        		<TD><%=Lang("Entered_By")%>:</TD>
        		<TD><%=objCase.EnteredBy.UserName%></TD>
        	</TR>
          <TR>
            <TD class="lhd_Required"><%=Lang("Title")%>:</TD>
            <%
            If blnPrintView = False Then
            %>

  	        	<%
				      If Application("ENABLE_ATTACHMENTS") = 1 Then
				      %>
              <TD colspan=3><INPUT id="Text2" style="WIDTH: 100%" name=tbxTitle value="<%=objCase.Title%>" readonly></TD>
	        	  <TD valign=top align=right>
                <INPUT style="WIDTH: 150px; BACKGROUND-COLOR: white" type="button" value="Attachments" name=btnAttachments id="Button1" onClick="VBScript:LoadAttachmentForm()" disabled>
              </TD>
  				    <%
  				    Else
  				    %>
              <TD colspan=4><INPUT id="Text3" style="WIDTH: 100%" name=tbxTitle value="<%=objCase.Title%>" readonly></TD>
              <%
  				    End If
    				  %>

            <%
            Else
            %>

              <TD colspan=3><%=objCase.Title%></TD>
              <td></td>

            <%
            End If
            %>
          </TR>
					<%
					Set rstAttachments = Server.CreateObject("ADODB.Recordset")
							
					rstAttachments.Open "SELECT * FROM tblFiles WHERE CaseFK='" & lngCaseID & "'", cnnDB
							
					If rstAttachments.BOF And rstAttachments.EOF Then
							
						' No records found
								
					Else
							
					  strHTML = ""
							
						Do 

              If Len(strHTML) > 0 Then
  							strHTML = strHTML & ",&nbsp;<A href=""FileView.asp?ID=" & rstAttachments.Fields("FilePK") & """ target=""_blank"">" &  rstAttachments.Fields("FileName") & "&nbsp;(" & CStr(Int(rstAttachments.Fields("FileSize")/1024)) & "k)" & "</A>"
  					  Else
  							strHTML = strHTML & "<A href=""FileView.asp?ID=" & rstAttachments.Fields("FilePK") & """ target=""_blank"">" &  rstAttachments.Fields("FileName") & "&nbsp;(" & CStr(Int(rstAttachments.Fields("FileSize")/1024)) & "k)" & "</A>"
  					  End If
							rstAttachments.MoveNext

						Loop Until rstAttachments.EOF

            %>								
            <TR>
              <TD></TD>
              <TD colspan="4">
                <FONT style="FONT-SIZE: 8pt">
                  <%
					  			Response.Write strHTML
					        %>
					      </FONT>
              </TD>
            </TR>
            <%							

					End If
							
					rstAttachments.Close
					Set rstAttachments = Nothing
					%>
          <TR>
            <TD vAlign=top><%=Lang("Detailed_Description")%>:</TD>
            <%
            If blnPrintView = False Then
            %>
              <TD colspan=4><TEXTAREA id=txtDescription style="SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-TRACK-COLOR: white; WIDTH: 100%; HEIGHT: 140px" name="txtDescription" readonly><%=objCase.Description%></TEXTAREA></TD>
            <%
            Else
            %>
              <TD colspan=4 vAlign=top><p><%=objCase.Description%></p></TD>
            <%
            End If
            %>
          </TR>
          <%

        		' Need to load all of the notes for this case
        		
        		strHTML = ""
        		
        		If objCase.CaseNotes.BOF And objCase.CaseNotes.EOF Then
        		
        			' No case notes to list
        			
        		Else
        	%>	
          <TR>
            <TD colspan="5"></TD>
          </TR>
          <TR>
            <TD vAlign=top><%=Lang("Notes")%>:</TD>
            <TD colspan=4>
          <%
        			While Not objCase.CaseNotes.EOF
        			
        				If objCase.CaseNotes.Item.IsPrivate Then
        				
        					If blnTECH = True or blnADMIN Then
        					
        						strHTML = strHTML & "<FONT Size=1><B>[" & DisplayDateTime(objCase.CaseNotes.Item.AddDate) & "&nbsp;&nbsp;" & objCase.CaseNotes.Item.Owner.UserName & " - PRIVATE]</B><BR>"
        						strHTML = strHTML & Replace(objCase.CaseNotes.Item.Note, Chr(13) & Chr(10), "<BR>") & "<BR></FONT>"
        					
        					Else
        					
        						' Do nothing, as the user isn't allow to see this note
        					
        					End If
        				
        				Else
        				
        					strHTML = strHTML & "<FONT Size=1><B>[" & DisplayDateTime(objCase.CaseNotes.Item.AddDate) & "&nbsp;&nbsp;" & objCase.CaseNotes.Item.Owner.UserName & "]</B><BR>"
        					strHTML = strHTML & Replace(objCase.CaseNotes.Item.Note, Chr(13) & Chr(10), "<BR>") & "<BR></FONT>"
        				
        				End If

        				objCase.CaseNotes.MoveNext
        			
        			WEnd
        			
          		Response.Write strHTML
        	%>	
            </TD>
          </TR>
          <%
        		End If
       		%>
        	<TR>
        		<TD colspan="5"></TD>
        	</TR>
        	<TR class="lhd_Heading2">
        		<TD colspan=5><%=Lang("Resolution")%></TD>
        	</TR>
        	<TR>
        	  <TD vAlign=top><%=Lang("Resolution")%>:</TD>
        	  <%
        	  If blnPrintView = False Then
        	  %>
        	    <TD colspan=4><TEXTAREA id=txtResolution style="SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-TRACK-COLOR: white; WIDTH: 100%; HEIGHT: 140px" name=txtResolution readonly><%=objCase.Resolution%></TEXTAREA></TD>
        	  <%
        	  Else
        	  %>
        	    <TD colspan=4 vAlign=top><p><%=objCase.Resolution%></p></TD>
        	  <%
        	  End If
        	  %>
        	</TR>
        	<TR>
        		<TD colspan="5"></TD>
        	</TR>
        	<TR>
        		<%
        		If blnPrintView = False Then
        		%>
        		  <td></td>
        		  <TD align=Left><a href="caseModify.asp?ID=<%=lngCaseID%>&PrintView=1"><%=Lang("Printer_Friendly_View")%></a></TD>
        		<%
        		Else
        		%>
        		  <TD colspan=2></TD>  
        		<%
        	  End If
        		%>
        		<%
        		If PERM_REOPEN_CASES = (PERM_REOPEN_CASES And binUserPermMask) AND blnCanModify = True AND blnPrintView = False Then
        		%>
        		  <TD colspan=3 align=right><INPUT style="WIDTH: 80px; BACKGROUND-COLOR: white" type="submit" value="<%=Lang("ReOpen")%>" name=btnReOpen id=btnReOpen></TD>
        		<%
        		Else
        		%>
        		  <TD colspan=3></TD>  
        		<%
        	  End If
        		%>
        	</TR>

        </TABLE>

      <%
      End If
      %>

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

Set objCase = Nothing
	
cnnDB.Close
Set cnnDB = Nothing

%>

