<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: admAssignment.asp
'  Date:     $Date: 2004/03/10 08:21:12 $
'  Version:  $Revision: 1.6 $
'  Purpose:  Administration page for creating/modifing assignments
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

<HTML>
<HEAD>

<META content="MSHTML 6.00.2600.0" name=GENERATOR></HEAD>

<LINK rel="stylesheet" type="text/css" href="Default.css">
<%
	Dim cnnDB
	Dim binUserPermMask, binRequiredPerm
  Dim rstR
  Dim blnSave, blnIsActive, blnMatch
  Dim objAssignment, objCollection, objCaseTypes, objCategories, objListItems
  Dim strMode, strIsActiveHTML, strSQL, strHTML, strHeading
  Dim intAssignmentID, intRepID, intCatID, intCaseTypeID
  Dim intLastUpdateByID, intUserID, intLastGroupID, intLastCaseTypeID
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
	
		intLastCaseTypeID = Empty

		While Not rstR.EOF
	
			If rstR.Fields("CaseTypePK") <> intLastCaseTypeID Then
			
				If Not IsEmpty(intLastCaseTypeID) Then
					strHTML = strHTML & "	</CASETYPE>" & Chr(13)
				Else
					' Do nothing
				End If
				
				strHTML = strHTML & "	<CASETYPE CaseTypeID=" & Chr(34) & rstR.Fields("CaseTypePK") & Chr(34) & ">" & Chr(13)
				intLastCaseTypeID = rstR.Fields("CaseTypePK")
						
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

	If Not IsEmpty(intLastCaseTypeID)Then
		strHTML = strHTML & "	</CASETYPE>" & Chr(13)
	Else
		' Do nothing
	End If

	strHTML = strHTML & "</CASETYPES_REPS>" & Chr(13)
	strHTML = strHTML & "</XML>"
	
	Response.Write strHTML


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
	
		intLastCaseTypeID = Empty

		While Not rstR.EOF
	
			If rstR.Fields("CaseTypePK") <> intLastCaseTypeID Then
			
				If Not IsEmpty(intLastCaseTypeID) Then
					strHTML = strHTML & "	</CASETYPE>" & Chr(13)
				Else
					' Do nothing
				End If

				strHTML = strHTML & "	<CASETYPE CaseTypeID=" & Chr(34) & rstR.Fields("CaseTypePK") & Chr(34) & ">" & Chr(13)
				intLastCaseTypeID = rstR.Fields("CaseTypePK")
				
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

	If Not IsEmpty(intLastCaseTypeID)Then
		strHTML = strHTML & "	</CASETYPE>" & Chr(13)
	End If

	strHTML = strHTML & "</CASETYPES_CATEGORIES>" & Chr(13)
	strHTML = strHTML & "</XML>" & Chr(13)
	
	Response.Write strHTML


%>

<SCRIPT language="VBScript">

	Sub ListCategoriesAndReps()
	
		Dim XML
		Dim xmlNode, xmlNodes
		Dim objCategoryList, objRepList
	

		' Generate list of associated Categories
	
		Set XML = Document.All("CASETYPES_CATEGORIES")
		Set xmlNodes = XML.SelectNodes("CASETYPES_CATEGORIES/CASETYPE[@CaseTypeID='" & Document.All.cbxCaseType.Value & "']/CATEGORY")

		Set objCategoryList = Document.All("cbxCat")

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
		objRepList.Options(0).InnerText = "(Unassigned)"

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

		intAssignmentID = Cint(Request.Form("tbxAssignmentID"))
		intCaseTypeID = Cint(Request.Form("cbxCaseType"))
		intCatID = Cint(Request.Form("cbxCat"))
		intRepID = CInt(Request.Form("cbxRep"))

		If Request.Form("chkIsActive") = "on" Then
			blnIsActive = True
		Else
			blnIsActive = lhd_False
		End If

		dteLastUpdate = SQLDateFormat(Day(Now()), Month(Now()), Year(Now())) & " " & FormatDateTime(Now(), vbShortTime)
		intLastUpdateByID = intUserID

		
		' Check for required fields
		
		If intCaseTypeID = 0 Or IsEmpty(intCaseTypeID) Then
		  DisplayError 1, "Case Type"	
		Else
      ' Do nothing		
		End If
		
		If intCatID = 0 Or IsEmpty(intCatID) Then
		  DisplayError 1, "Category"	
		Else
      ' Do nothing				
		End If
		
		If intRepID = 0 Or IsEmpty(intRepID) Then
		  DisplayError 1, "Assigned Rep"	
		Else
      ' Do nothing				
		End If


		' Now save/update the Assignment details

		Set objAssignment = New clsAssignment

		objAssignment.ID = intAssignmentID

		If Not objAssignment.Load Then
		
			' Assignment does not exist, thus we need to now check that we aren't creating a
			' duplicate assignment

		  Set objCollection = New clsCollection
  		
	    objCollection.CollectionType = objCollection.clAssignment
	    objCollection.Query = "SELECT * FROM tblAssignments"

	    If Not objCollection.Load Then
  	  			
	  	  ' Recordset failed to load
  	  				
	    Else
  	  									
	  	  If objCollection.BOF And objCollection.EOF Then
  		
		      ' Do nothing as not results returned
  		
		    Else
  		  
		      While Not objCollection.EOF And Not blnMatch
  		    
		        If objCollection.Item.CaseTypeID = intCaseTypeID And objCollection.Item.CatID = intCatID Then
		          blnMatch = True
		        Else
		          blnMatch = False
		        End If
  		    
		        objCollection.MoveNext
  		    
		      WEnd
  		    
		      If blnMatch Then
		        DisplayError 3, "You can not create a duplicate assignment."
		      Else
		        ' Do nothing
		      End If
  		  
		    End If
  		  
		  End If
  		
		  Set objCollection = Nothing

		Else
		
			' Assignment exists, now we need to check that any changes we are making don't
			' conflict with already existing assignments
		
		  Set objCollection = New clsCollection
  		
	    objCollection.CollectionType = objCollection.clAssignment
	    objCollection.Query = "SELECT * FROM tblAssignments"

	    If Not objCollection.Load Then
  	  			
	  	  ' Recordset failed to load
  	  				
	    Else
  	  									
	  	  If objCollection.BOF And objCollection.EOF Then
  		
		      ' Do nothing as not results returned
  		
		    Else
  		  
		      While Not objCollection.EOF And Not blnMatch
  		    
		        If objCollection.Item.CaseTypeID = intCaseTypeID And objCollection.Item.CatID = intCatID And objCollection.Item.ID <> intAssignmentID Then
		          blnMatch = True
		        Else
		          blnMatch = False
		        End If
  		    
		        objCollection.MoveNext
  		    
		      WEnd
  		    
		      If blnMatch Then
		        DisplayError 3, "You can not create a duplicate assignment."
		      Else
		        ' Do nothing
		      End If
  		  
		    End If
  		  
		  End If
  		
		  Set objCollection = Nothing

		End If

		' Check the the fields and leave Null if nothing is set.

		objAssignment.CaseTypeID = intCaseTypeID
		objAssignment.CatID = intCatID
		objAssignment.RepID = intRepID
		objAssignment.IsActive = blnIsActive
		objAssignment.LastUpdate = dteLastUpdate
		objAssignment.LastUpdateByID = intLastUpdateByID

						
		If Not objAssignment.Update Then
						
			' Failed to create/save Assignment
							
		Else
						
			intAssignmentID = objAssignment.ID
			strHeading = Lang("Assignment Saved")
						
		End If
						
		Set objAssignment = Nothing
		%>
		<TR>
			<TD>
				<TABLE class=Normal width="100%" border=0 cellSpacing=0 cellPadding=1>
					<TR class="lhd_Heading1">
						<TD colspan=5 align=center><%=strHeading%></TD>
					</TR>
		      <TR>
		        <TD align=right colspan=5><a style="FONT-SIZE: 8pt" href="admAssignmentList.asp?Page=1"><%=Lang("Manage_Assignments")%></a></TD>
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
						<TD colspan=3 align="Left">Assignment changes have been successfully saved.</TD>
						<TD></TD>
					</tr>
					<tr>
					  <td colspan=5></td>
					</tr>
				</TABLE>
			</TD>
		</TR>
	<%
	Else
	
		' Mode: 1 - To create a new Assignment
		'		2 - Edit a Assignment
	
'		strMode = Request.QueryString("mode")
	
		If Request.QueryString.Count = 0 Then
	
			' Create a new record
	
			strMode = 1
			intAssignmentID = 0
			strHeading = Lang("New_Assignment")
			
		Else
	
			' Edit a record determine by the Assignment ID passed via the QueryString
	
			strMode = 2
			intAssignmentID = Request.QueryString("id")
			strHeading = Lang("Modify_Assignment")
			
		End If
			

		Select Case strMode
		
			Case 1	' Create new Assignment
			
				intCaseTypeID = 0
				intCatID = 0
				intRepID = 0
				
				blnIsActive = lhd_True
				strIsActiveHTML = "CHECKED"

				dteLastUpdate = ""
				intLastUpdateByID = 0


			Case 2  ' Edit Assignment

				' Get the Assignment ID we want to edit and load the record

				Set objAssignment = New clsAssignment

				objAssignment.ID = intAssignmentID
			
				If Not objAssignment.Load Then
				
					' Couldn't load user for some reason
				
				Else
				
					intCaseTypeID = objAssignment.CaseTypeID
					intCatID = objAssignment.CatID
					intRepID = objAssignment.RepID
					
					If objAssignment.IsActive = True Then
						strIsActiveHTML = "CHECKED"
					Else
						strIsActiveHTML = ""
					End If
					
					dteLastUpdate = objAssignment.LastUpdate
					intLastUpdateByID = objAssignment.LastUpdateByID

				End If

				Set objAssignment = Nothing
				
			Case Else
				' Do nothing
				
		End Select

		%>

		<TR>
			<TD>
				<TABLE class=Normal cellSpacing=0 cellPadding=1 width="100%" border=0 bgColor=white>
					<FORM action="admAssignment.asp" method="post" id=frmAssignment name=frmAssignment>
					<INPUT id=tbxAssignmentID name=tbxAssignmentID type=hidden value="<%=intAssignmentID%>">
					<INPUT id=tbxSave name=tbxSave type=hidden value="1">
					<TR class="lhd_Heading1">
						<TD colspan=5 align=center><%=strHeading%></TD>
					</TR>
		      <TR>
		        <TD align=right colspan=5><a style="FONT-SIZE: 8pt" href="admAssignmentList.asp?Page=1"><%=Lang("Manage_Assignments")%></a></TD>
		      </TR>
					<TR>
						<TD width="22%" class="lhd_Required"><%=Lang("Case_Type")%>:</TD>
						<TD width="25%">
							<SELECT id=cbxCaseType name=cbxCaseType style="WIDTH: 100%" onChange="VBScript:ListCategoriesAndReps()">
								<OPTION VALUE="0" SELECTED>(None)</OPTION>
								<%
								Set objCollection = New clsCollection
					    
								objCollection.CollectionType = objCollection.clCaseType
								objCollection.Query = "SELECT * FROM tblCaseTypes " &_
													  "WHERE IsActive=" & lhd_True & " " &_
													  "ORDER BY CaseTypeOrder ASC"
					    
								If Not objCollection.Load Then
					    
									' Didn't load
									
								Else
					    
								    If objCollection.BOF And objCollection.EOF Then
					    
										' No records returned
										
									Else
									
										Do While Not objCollection.EOF

											If objCollection.Item.ID = intCaseTypeID Then
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
						<TD class="lhd_Required"><%=Lang("Category")%>:</TD>
						<TD>
							<SELECT id=cbxCat name=cbxCat style="WIDTH: 100%">
								<OPTION VALUE="0">(None)</OPTION>
								<%
								If Not IsNull(intCaseTypeID) Then
		    
									Set objCollection = New clsCollection
												    
									objCollection.CollectionType = objCollection.clCategory
									objCollection.Query = "SELECT * FROM tblCategories " &_
														  "WHERE IsActive=" & lhd_True & " AND CaseTypeFK=" & intCaseTypeID & " " &_
														  "ORDER BY CatOrder ASC"
									
												    
									If Not objCollection.Load Then
												    
										' Didn't load
														
									Else
												    
									    If objCollection.BOF And objCollection.EOF Then
												    
											' No records returned
															
										Else
														
											Do While Not objCollection.EOF

												If objCollection.Item.ID = intCatID Then
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
									
									Set objCollection = Nothing
									
								Else
			
									' Do Nothing
									
								End If
								%>
							</SELECT>
						</TD>
						<TD></TD>
						<TD></TD>
						<TD></TD>
					</TR>
						<TD><B><%=Lang("Assigned_Rep")%>:</B></TD>
						<TD>
							<SELECT id=cbxRep name=cbxRep style="WIDTH: 100%">
								<OPTION selected value="0">(Unassigned)</OPTION>
								<%
								If Not IsNull(intCaseTypeID) Then
            
									Set objCollection = New clsCollection
												    
									objCollection.CollectionType = objCollection.clContact
									objCollection.Query = "SELECT tblContacts.* FROM tblContacts " &_
														  "INNER JOIN (tblCaseTypes INNER JOIN (tblGroups INNER JOIN tblGroupMembers ON tblGroups.GroupPK = tblGroupMembers.GroupFK) ON tblCaseTypes.RepGroupFK = tblGroups.GroupPK) ON tblContacts.ContactPK = tblGroupMembers.ContactFK " &_
														  "WHERE tblCaseTypes.IsActive=" & lhd_True & " AND tblCaseTypes.CaseTypePK=" & intCaseTypeID & " AND tblContacts.IsActive=" & lhd_True & " AND tblGroups.IsActive=" & lhd_True & " " &_
														  "ORDER BY tblContacts.UserName ASC"
												    
									If Not objCollection.Load Then
												    
										' Didn't load
														
									Else
												    
									    If objCollection.BOF And objCollection.EOF Then
												    
											' No records returned
															
										Else
														
											Do While Not objCollection.EOF

												If objCollection.Item.ID = intRepID Then
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
			
								Else
			
									' Do nothing
			
								End If
								%>
							</SELECT>
						</TD>
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
			</TD>
		</TR>
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
