<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: caseNewPost.asp
'  Date:     $Date: 2004/03/23 05:37:27 $
'  Version:  $Revision: 1.7 $
'  Purpose:  Used the post the user entered information to the Case table
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
	
	Dim binUserPermMask
	
	Dim blnNotifyUser, blnPrivateNote, blnIsActive, blnIsMember
	
	Dim rstAttachments
	
	Dim objAssignmentMethods, objAssignments, objGroup, objNote, objCase, objCaseType, objContact
	
	Dim lngCaseID
	
	Dim intUserID, intStatusID, intRepID, intEnteredByID, intLastUpdateByID, intCaseTypeGroupID
	Dim intContactID, intCaseTypeID, intCategoryID, intGroupID, intPriorityID, intBillTypeID, intDeptID
	
	Dim strHTML, strSQL, strAltEmail, strCc, strTemporaryCaseID
	Dim strTitle, strDescription, strNote, strResolution, strMinutesSpent
	
	Dim dteRaisedDate, dteClosedDate, dteLastUpdate, dteAddDate
	
	

	' Setup connection to database and get page values.

	Set cnnDB = CreateConnection

	intUserID = GetUserID
	binUserPermMask = GetUserPermMask
	

	' Get Case details from Form

	If CInt(Request.Form("cbxContact")) = 0 Then
	
		intContactID = CInt(Request.Form("tbxContact"))
  	intDeptID = CInt(Request.Form("tbxDeptID"))
		
	Else
	
		intContactID = CInt(Request.Form("cbxContact"))
		
		Set objContact = new clsContact
		
		objContact.ID = intContactID
		
		If Not objContact.Load Then
		  ' Contact didn't load
		Else
		  intDeptID = objContact.DeptID
		End If
		
		Set objContact = Nothing
		
	End If

	intEnteredByID = CInt(Request.Form("tbxEnteredByID"))
	intCaseTypeID = CInt(Request.Form("cbxCaseType"))
	intCategoryID = CInt(Request.Form("cbxCategory"))
	intPriorityID = CInt(Request.Form("cbxPriority"))
	intStatusID = CInt(Request.Form("cbxStatus"))
	intRepID = CInt(Request.Form("cbxRep"))
	intBillTypeID = CInt(Request.Form("cbxBillType"))

	strAltEmail = Request.Form("tbxAltEmail")
	strTitle = Request.Form("tbxTitle")
	strDescription = Request.Form("txtDescription")
	strResolution = Request.Form("txtResolution")
	strNote = Request.Form("txtNotes")
	strMinutesSpent = Request.Form("tbxMinutesSpent")
	strCc = Request.Form("tbxCc")
	strTemporaryCaseID = Request.Form("tbxCaseID")

	dteRaisedDate = SQLDateFormat(Day(Now()), Month(Now()), Year(Now())) & " " & FormatDateTime(Now(), vbShortTime)
	dteClosedDate = ""
	dteLastUpdate = SQLDateFormat(Day(Now()), Month(Now()), Year(Now())) & " " & FormatDateTime(Now(), vbShortTime)
	intLastUpdateByID = intUserID


' -------------------------------------------------------------------------------
' Check required fields

If intCaseTypeID = 0 Then
	DisplayError 1, "Case Type"	
Else
	' Do Nothing
End If

If intCategoryID = 0 Then
	DisplayError 1, "Category"
Else
	' Do Nothing	
End If

If Len(strTitle) = 0 Then
	DisplayError 1, "Title"
Else
	' Do Nothing
End If
	

' -------------------------------------------------------------------------------

	' We need to try and assign this case if it hasn't already been assigned

	' Automatically assign the case
		
	'  How do we know what field on the Case to use for the assignment. We need a set the ASSIGN_METHODS
	' so that we know which one we are useing so that we get the right data from the right field.	
		
	' Note - Another way we could do this would be the name the AssignMethods the name of the field
	' we use on the form then the query could be:
	' ( ... " AND AssignValueFK=(SELECT AssignMethod FROM tblAssignMethods WHERE AssignMethodPK=" & CStr(Application("ASSIGN_METHOD"))
	' NOTE !!! - This is how I have implemented it below making sure the AssignMethod name can be used map
	' the required variable


	' -------------------------------------------------------------------------------
	' Rep Assignments


	' First determine the default Rep Group assigned to this Case Type
	
	Set objCaseType = New clsCaseType
				
	objCaseType.ID = intCaseTypeID
				
	If Not objCaseType.Load Then

		' This occurance is not likely as RepGroupFk is a required field

	Else

		intCaseTypeGroupID = objCaseType.RepGroupID
		
	End If
				
	Set objCaseType = Nothing
	

	' Check if Case was manually assigned
	
	If IntRepID > 0 Then
	
		' Need to determine if the person enterind the case is trying to assign it
		' to when he doesn't belong to the CaseType's RepGroupFK for the Case
		' submitted.
		
		Set objGroup = New clsGroup
		
		objGroup.ID = intCaseTypeGroupID
			
		If Not objGroup.Load Then
		
			' This occurance is not likely as RepGroupFk is a required field
			
			intGroupID = Empty
		
		Else

			If Not objGroup.IsMember(intUserID) Then

				' Current user isn't a member of the Case Type Rep Group and hence
				' can't manually assign the case. So now we need to automatically assign 
				' the Case
			
				intRepID = Empty
				intGroupID = Empty
			
			Else

				' User is a member of the Rep Group assigned to this Case Type and hence
				' is allowed to manually assign a Rep.
				
				intGroupID = intCaseTypeGroupID
			
			End If		
			
		End If
			
		Set objGroup = Nothing
		
	Else
	
		' Do Nothing.  So now we need to automatically assign the Case
		
		intRepID = Empty
		intGroupID = Empty
		
		
	End If
	
	
	' Check if we need to go through the automatic assignment process, determine by
	' if a Rep was legally selected
	
	If IsEmpty(intRepID) And IsEmpty(intGroupID) Then
	
		' Case not manually assigned.  Look up the tblAssignments to determine who the 
		' the case should be assigned to
		
		Set objAssignments = New clsCollection

		objAssignments.CollectionType = objAssignments.clAssignment
		objAssignments.Query = "SELECT * FROM tblAssignments " &_
							   "WHERE IsActive=" & lhd_True & " " &_
									  "AND CaseTypeFK=" & intCaseTypeID & " " &_
									  "AND CatFK=" & intCategoryID
										    
		If Not objAssignments.Load Then
										    
			' There has been no assignment method set. We now default to the RepGroupFK
			' allocated to the Case Type in the tblCaseTypes table.
			
			intRepID = Empty
			intGroupID = intCaseTypeGroupID
												
		Else

			If objAssignments.BOF And objAssignments.EOF Then

				' There has been no assignment method set. We now default to the RepGroupFK
				' allocated to the Case Type in the tblCaseTypes table.
			
				intRepID = Empty
				intGroupID = intCaseTypeGroupID

			Else
			
				' Assign Case a Rep and Group

				intRepID = objAssignments.Item.RepID
				intGroupID = intCaseTypeGroupID
			
			End If
	
		End If

		Set objAssignments = Nothing

	Else
	
		' Do nothing as case has been manually assigned
	
	End If


	' -------------------------------------------------------------------------------
	' Save Case Details
	    
	Set objCase = New clsCase
		
	With objCase

		.ContactID = intContactID
		.RepID = intRepID
		.GroupID = intGroupID
		.StatusID = intStatusID
		.CatID = intCategoryID
		.PriorityID = intPriorityID
		.CaseTypeID = intCaseTypeID
		.Title = strTitle
		.Description = strDescription
		.Resolution = strResolution
		.AltEMail = strAltEmail
		.Cc = strCc
		.DeptID = intDeptID

		If Len(dteRaisedDate) > 0 Then
			.RaisedDate = CDate(dteRaisedDate)
		End If
		
		If Len(dteClosedDate) > 0 Then
			.ClosedDate = CDate(dteClosedDate)
		End If
		
		.EnteredByID = intEnteredByID
		.IsActive = True

		If Len(dteLastUpdate) > 0 Then
			.LastUpdate = CDate(dteLastUpdate)
		End If

		.LastUpdateByID = intLastUpdateByID

	End With


	' Write the Case data to the database

	If Not objCase.Update Then
		' Raise Error, Case not updated
	Else
		lngCaseID = objCase.ID
	End If

	Set objCase = Nothing


	'  Now that we have created the Case and obtained a Case ID we can add any
	' additional notes to the tblNotes table
	
	If strNote <> "" Or strMinutesSpent <> "" Then
	
		If Request.Form("chkPrivateNote") = "on" Then
		   blnPrivateNote = lhd_True
		Else
		   blnPrivateNote = lhd_False
		End If

		dteAddDate = dteLastUpdate
	
		Set objNote = New clsNote
		
		With objNote
			.CaseID = lngCaseID
			.OwnerID = intUserID
'			.Note = strNote & "  (" & strMinutesSpent & " Minutes Logged)"
			.Note = strNote & "<BR>"
			.MinutesSpent = strMinutesSpent
			If Not IsEmpty(dteAddDate) Then
				.AddDate = CDate(dteAddDate)
			Else
				' Do nothing
			End If
			.BillTypeID = intBillTypeID
			.IsPrivate = blnPrivateNote
			If Not IsEmpty(dteLastUpdate) Then
				.LastUpdate = CDate(dteLastUpdate)
			Else
				' Do nothing
			End If
			.LastUpdateByID = intLastUpdateByID
		End With
		  
		If Not objNote.Update Then
			Response.Write objNote.LastError	' Case note update failed
		Else
			' Case note Updated
		End If

		Set objNote = Nothing
		
	Else
	
		' No notes to be recorded
		
	End If
	
	
	' -------------------------------------------------------------------------------
	' Attachments
	
	' Next thing is to match and link any attached files to this particular case.  This
	' means to replace the temporary CaseID with the newly generate CaseID in the 
	' table tblFiles.  (Note:  This is only done for newly created cases)
	
	Set rstAttachments = Server.CreateObject("ADODB.Recordset")
							
	rstAttachments.Open "SELECT * FROM tblFiles WHERE CaseFK='" & strTemporaryCaseID & "'", cnnDB
							
	If rstAttachments.BOF And rstAttachments.EOF Then
							
		' No records found
								
	Else
							
		cnnDB.Execute "UPDATE tblFiles SET CaseFK='" & CStr(lngCaseID) & "' WHERE CaseFK='" & strTemporaryCaseID & "'", , adExecuteNoRecords

	End If
							
	rstAttachments.Close
	
	Set rstAttachments = Nothing

	
	' -------------------------------------------------------------------------------
	' EMail Notifications

	'  Now all Case information has been saved, lets send out the appropriate emails

	If Application("ENABLE_EMAIL") = 1 Then

		' NOTES
		' -----
		' 1.) Send an email to the assigned Technician, if any.  However do not send an email
		' to the technician if the assigned technication is the one who has entered in the
		' case
		'
		' 2.)  Send an email the unassigned email address if not Technician was assign automatically
		' or otherwise ( I.e. intRepID is Empty)
		
		' 3.)  Send an email to the contact if the case was raised on behalf of that contact
		' Do not send an email to the Contact if he/she raised it.
		'
		' 4.)  An overiding flag can be used to force an email to be sent to the Requestor

		If Not IsEmpty(intRepID) Then

			' Case was manually/automatically assigned so we need to send an email
			
			SendEmail lngCaseID, "CASE_ASSIGNED_REP", "REP"

		Else
		
			' No Rep has been assigned so we will send email to the entire assigned
			' group
			
			If Not IsEmpty(intGroupID) Then
			
			    SendEmail lngCaseID, "CASE_UNASSIGNED_REP", "GROUP"
			
			Else
			
				' No assign group so we should send this to the system administrator
				' which is set in tblParameters table

			    SendEmail lngCaseID, "CASE_UNASSIGNED_ADMIN", "ADMIN"
			
			End If

		End If


		' Finally send an email confirmation that their case has been submitted
		
		SendEmail lngCaseID, "CASE_SUBMITTED_USER", "USER"

	Else
	
		' No nothing, as email notfications haven't been enabled
	
	End If
	
%>	    

<HTML>

<HEAD>
	
	<META content="MstrHTML 6.00.2600.0" name=GENERATOR>
</HEAD>

<LINK rel="stylesheet" type="text/css" href="Default.css">

<BODY>
<P align=center>
<TABLE class=Normal align=center style="WIDTH: 680px" cellSpacing=1 cellPadding=1 width="680" border=0>
	<TR>
		<TD>
			<%
			Response.Write DisplayHeader
			%>
		</TD>
	</TR>
	<TR>
		<TD>
			<TABLE class="lhd_Box" cellSpacing=0 cellPadding=1 width="100%" border=0 bgColor=white>
				<TR class="lhd_Heading1">
					<TD colspan=5><%=Lang("Case_Submitted")%></TD>
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
					<TD colspan=3>
						<B><%=Lang("Case")%> #<%=lngCaseID%></B>
						<BR><BR>
						Thank you, your request had been received and will be actioned as soon as possible.
						You can view your new request by clicking <B><A href="caseModify.asp?ID=<%=lngCaseID%>">here</A></B>.
						<BR><BR>
						The progress of all your active requests can be viewed and monitored via the Main Menu.
					</TD>
					<TD></TD>
				</TR>
				<TR>
					<TD colspan=5></TD>
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
</TABLE>
</P>
</BODY>
</HTML>

<%

cnnDB.Close
Set cnnDB = Nothing

%>