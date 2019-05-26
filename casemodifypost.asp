<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: caseModiyPost.asp
'  Date:     $Date: 2004/03/23 05:37:26 $
'  Version:  $Revision: 1.7 $
'  Purpose:  Used the post and modification made to a case
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


' Declare variables

Dim cnnDB

Dim binUserPermMask, binRequiredPerm

Dim blnStatusChanged, blnUserEmailSent, blnReOpen
Dim blnNotifyUser, blnPrivateNote, blnUpdate, blnRepChanged, blnCanModify

Dim intContactID, intCaseTypeID, intCategoryID, intGroupID, intPriorityID, intDeptID
Dim intBillTypeID, intStatusID, intRepID, intEnteredByID, intUserID, intLastUpdateByID, intCaseTypeGroupID

Dim lngCaseID

Dim strAltEMail, strTitle, strDescription, strNote, strResolution, strMinutesSpent
Dim strSQL, strHTML, strCc

Dim dteClosedDate, dteLastUpdate, dteAddDate

Dim objNote, objCase, objContact, objListItem, objCategory, objCaseType
Dim objCollection, objAssignments, objGroup



' Setup connection to database and get user details

Set cnnDB = CreateConnection

intUserID = GetUserID
binUserPermMask = GetUserPermMask


lngCaseID = CLng(Request.Form("tbxCaseID"))
  
intContactID = CInt(Request.Form("tbxContactID"))
intRepID = CInt(Request.Form("cbxRep"))
intGroupID = CInt(Request.Form("tbxGroup"))
intStatusID = CInt(Request.Form("cbxStatus"))
intCategoryID = CInt(Request.Form("cbxCategory"))
intPriorityID = CInt(Request.Form("cbxPriority"))
intCaseTypeID = CInt(Request.Form("cbxCaseType"))
'intDeptID = CInt(Request.Form("tbxDeptID"))

strTitle = Request.Form("tbxTitle")
strDescription = Request.Form("txtDescription")
strResolution = Request.Form("txtResolution")
strCc = Request.Form("tbxCc")
strAltEMail = Request.Form("tbxAltEmail")


If Request.Form("chkNotifyUser") = "on" Then
	blnNotifyUser = True
Else
	blnNotifyUser = False
End If

'		If Request.Form("chkIsActive") = "on" Then
'			blnIsActive = 1
'		Else
'			blnIsActive = 0
'		End If
'		dteRaisedDate = Request.Form("tbxRaisedDate"))
'		intEnteredByID = CInt(Request.Form("tbxEnteredByID"))

strMinutesSpent = Request.Form("tbxMinutesSpent")

' Update case details, but first load the existing data
		    
Set objCase = New clsCase

objCase.ID = lngCaseID
		
If Not objCase.Load Then
	' Raise Error, case didn't load
Else
	' Case loaded
End If

'Check to see if user has permission to modify the case.

blnCanModify = CanModifyCase(objCase,intUserID,binUserPermMask)

If Not blnCanModify Then
	DisplayError 4, "This User does not have permission to modify this case.  "
End If
			

If CInt(Request.Form("tbxReOpen")) = 1 Then

  ' We will reopen the case, i.e. set the status back to open
  blnReOpen = True

  intContactID = objCase.ContactID
  intRepID = objCase.RepID
  intGroupID = objCase.GroupID
  intStatusID = Application("DEFAULT_STATUS")
  intCaseTypeID = objCase.CaseTypeID
  intCategoryID = objCase.CatID
  intPriorityID = objCase.PriorityID
  intDeptID = objCase.DeptID

  strCc = objCase.Cc
  strAltEMail = objCase.AltEmail
  strTitle = objCase.Title
  strDescription = objCase.Description
  strResolution = objCase.Resolution

Else

  ' Do nothing, as we are not reopening the case
  blnReOpen = False

End If


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


' -------------------------------------------------------------------------------
' Rep Assignments


' Only reprocess assignments if the CaseType has been changed

If objCase.CaseTypeID <> intCaseTypeID Then
	
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
	
	If intRepID > 0 Then
	
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

Else

	' Do nothing, as the Case Type has not changed
	
End If

' -------------------------------------------------------------------------------


'  We also need to include audit logging of any change of Priority, Status
' Assignment, Case Type and/or Category.
		
		
strNote = ""
		
If intPriorityID <> objCase.PriorityID Then

	Set objListItem = New clsListItem
		
	objListItem.ID = intPriorityID
		
	If Not objListItem.Load Then
	   Response.Write objListItem.LastError
	Else
     strNote = strNote & "PRIORITY: " & objCase.Priority.ItemName & " -> " & objListItem.ItemName  & "<BR>"
  End If
        
  Set objListItem = Nothing
        
Else
	
	' Do nothing
		
End If


If intStatusID <> objCase.StatusID Then

	blnStatusChanged = True
	
	Select Case intStatusID

		Case Application("STATUS_OPEN")		' STATUS_OPEN


		Case Application("STATUS_CLOSED")	' STATUS_CLOSED
			
			' Check for required fields on closing of a Case
				
			If Len(strResolution) > 0 Then
	
				' Resolution has been entered
				dteClosedDate = dteLastUpdate
	
			Else
	
				DisplayError 1, "Resolution"
	
			End If
			
			
		Case Else
			' Do nothing
		
	End Select

	Set objListItem = New clsListItem
		
	objListItem.ID = intStatusID
		
	If Not objListItem.Load Then
	   Response.Write objListItem.LastError
	Else
	   strNote = strNote & "STATUS: " & objCase.Status.ItemName & " -> " & objListItem.ItemName & "<BR>"
	End If
		
	Set objListItem = Nothing

Else

	blnStatusChanged = False

End If


If intCaseTypeID <> objCase.CaseTypeID Then
		
	Set objCaseType = New clsCaseType
		
	objCaseType.ID = intCaseTypeID
		
	If Not objCaseType.Load Then
	   Response.Write objCaseType.LastError
	Else
	   strNote = strNote & "CASE TYPE: " & objCase.CaseType.CaseTypeName & " -> " & objCaseType.CaseTypeName & "<BR>"
	End If
		
	Set objCaseType = Nothing
		
Else

	' Do nothing
	
End If

If intCategoryID <> objCase.CatID Then
	
	Set objCategory = New clsCategory
		
	objCategory.ID = intCategoryID
		
	If Not objCategory.Load Then
	   Response.Write objCategory.LastError
	Else
	   strNote = strNote & "CATEGORY: " & objCase.Cat.CatName & " -> " & objCategory.CatName & "<BR>"
	End If
		
	Set objCategory = Nothing
		
Else

	' Do nothing
	
End If


If (intRepID <> objCase.RepID) Or (Not IsEmpty(intRepID) And IsNull(objCase.RepID)) Or (IsEmpty(intRepID) And Not IsNull(objCase.RepID)) Then

	blnRepChanged = True

  If intRepID > 0 Then

	  Set objContact = New clsContact

	  objContact.ID = intRepID
    	
	  If Not objContact.Load Then
	    
	    ' List Item not found
	    
	  Else

      If objCase.RepID > 0 Then
  	    strNote = strNote & "ASSIGNMENT: " & objCase.Rep.UserName & " -> " & objContact.UserName & "<BR>"
  	  Else
  	    strNote = strNote & "ASSIGNMENT: Unassigned -> " & objContact.UserName & "<BR>"
  	  End If
  	  
	  End If
    	
	  Set objContact = Nothing
  	  
	Else

    If objCase.RepID > 0 Then
      strNote = strNote & "ASSIGNMENT: " & objCase.Rep.UserName & " -> Unassigned<BR>"
  	Else
  	  ' Do nothing
  	End If
  
	End If 
	  	
Else
	
	blnRepChanged = False
		
End If
		
' Save the audit information to tblNotes
		
If strNote <> "" Then

	' Save the notes
		
	dteAddDate = dteLastUpdate
			
	Set objNote = New clsNote
			
	With objNote
		.CaseID = lngCaseID
		.OwnerID = intUserID
		.Note = strNote
		.MinutesSpent = 0
		.IsPrivate = lhd_True
		If Not IsEmpty(dteAddDate) Then
			.AddDate = CDate(dteAddDate)
		Else
			' Do nothing
		End If
		.BillTypeID = intBillTypeID
		If Not IsEmpty(dteLastUpdate) Then
			.LastUpdate = CDate(dteLastUpdate)
		Else
			' Do nothing
		End If
		.LastUpdateByID = intLastUpdateByID
	End With

	If Not objNote.Update Then
		Response.Write objNote.LastError
	Else
		' Case note Updated
	End If
			
	Set objNote = Nothing
		
Else
		
	' No notes need to be saved
		
End If
		
'  Now we need to check if any notes have been added.

strNote = Request.Form("txtNotes")
		
If strNote <> "" Or strMinutesSpent <> "" Then

	If Request.Form("chkPrivateNote") = "on" Then
		blnPrivateNote = lhd_True
	Else
		blnPrivateNote = lhd_False
	End If

	' Save the notes
			
	dteAddDate = dteLastUpdate

	Set objNote = New clsNote
			
	With objNote
		.CaseID = lngCaseID
		.OwnerID = intUserID
		
		If Len(strNote) > 0 Then
		
		  If Len(strMinutesSpent) > 0 Then
  	  	.Note = "TIME SPENT: " & strMinutesSpent & " Minutes<BR>.<BR>" & strNote & "<BR>"
		  Else
  	  	.Note = strNote & "<BR>"
		  End If
		  
		Else
		
  	  .Note = "TIME SPENT: " & strMinutesSpent & " Minutes<BR>"
		
		End If
		
		.MinutesSpent = strMinutesSpent
		.IsPrivate = blnPrivateNote
		If Not IsEmpty(dteAddDate) Then
			.AddDate = CDate(dteAddDate)
		Else
			' Do nothing
		End If
		.BillTypeID = intBillTypeID
		If Not IsEmpty(dteLastUpdate) Then
			.LastUpdate = CDate(dteLastUpdate)
		Else
			' Do nothing
		End If
		.LastUpdateByID = intLastUpdateByID
	End With

	If Not objNote.Update Then
		Response.Write objNote.LastError
	Else
		' Case note Updated
	End If
			
	Set objNote = Nothing
		
Else
		
	' No notes need to be saved
		
End If

' Now update the Case details
	
With objCase

	.GroupID = intGroupID
	.RepID = intRepID
	.StatusID = intStatusID
	.CatID = intCategoryID
	.PriorityID = intPriorityID
	.CaseTypeID = intCaseTypeID
	.Title = strTitle
	.Description = strDescription
	.Resolution = strResolution
	.AltEMail = strAltEMail
	.Cc = strCc
	.DeptID = intDeptID

	If Not IsEmpty(dteClosedDate) Then
		.ClosedDate = CDate(dteClosedDate)
	Else
		' Do nothing
	End If

	If Not IsEmpty(dteLastUpdate) Then
		.LastUpdate = CDate(dteLastUpdate)
	Else
		' Do nothing
	End If

	.LastUpdateByID = intLastUpdateByID

'	Later on we may allow an Admin to update these fields
'
'	.ContactID = intContactID
'	.RaisedDate = dteRaisedDate
'	.IsActive = blnIsActive
'	.EnteredByID = intEnteredByID

End With

If Not objCase.Update Then
	Response.Write objCase.LastError
Else
	' Case Updated
End If
					

' -------------------------------------------------------------------------------
' EMail Notifications

'  Now all Case information has been saved, lets send out the appropriate emails

If Application("ENABLE_EMAIL") = 1 Then

	' NOTES
	' -----
	' 1.) Send a "Case Re-assigned" email to the assigned Technician, if the case was re-assigned.
	'	  However do not send an email to the technician if the assigned technication is the
	'	  one who has entered in the case
	'
	' 2.) Send a "Case Un-Assigned" email if the case has been change to have no assigned
	'	  Rep. ( I.e. intRepID is Empty)
	'
	' 3.) Send a "Case Closed" email to the Case contact (Requestor) if the case was
	'	  successfully closed
	'
	' 4.) Send an email to the Case contact if the "Notify User" flag has been checked
	'
	' 5.) Send an email to the assigned Rep, (or assigned Group), is the Contact (Requestor)
	'     updates the Case.
	
	blnUserEmailSent = False
	
	If blnRepChanged = True Then
	
		If IsEmpty(intRepID) Then
		
			' Send the "Case Un-Assigned" email
			SendEmail lngCaseID, "CASE_UNASSIGNED_REP", "GROUP"
		
		Else
		
			' Send the "Case Re-Assigned" email
			SendEmail lngCaseID, "CASE_REASSIGNED_REP", "REP"
			
		End If
	    
	Else
	
		' Do nothing
	
	End If


	If blnStatusChanged = True Then
	
		Select Case intStatusID
	
			Case Application("STATUS_REOPENED")		' STATUS_OPEN

				' Send the "Case Reopened" email
				SendEmail lngCaseID, "CASE_REOPENED_USER", "REP"


			Case Application("STATUS_CANCELLED")	' STATUS_CLOSED

				' Send the "Case Closed" email
				SendEmail lngCaseID, "CASE_CANCELLED_USER", "USER"
				blnUserEmailSent = True
				

			Case Application("STATUS_CLOSED")	' STATUS_CLOSED

				' Send the "Case Closed" email
				SendEmail lngCaseID, "CASE_CLOSED_USER", "USER"
				blnUserEmailSent = True
	

			Case Else
			
				' Do nothing
	
		End Select
		
	Else
	
		' Do nothing
	
	End If
	
	    
	If blnNotifyUser = True And blnUserEmailSent = False Then
	
		' Send an email to the Contact who raised the case
		SendEmail lngCaseID, "CASE_UPDATED_USER", "USER"

	Else
		
		' Do nothing
	
	End If
	    
	
	' Need to send an email to the assigned Rep or Group if the Contact (Requestor)
	' has updated the Case details

	If (intUserID = intContactID) Or InStr(1, strCc, Session("Username")) > 0 Then
	
		' Send an "Case Updated" email to the Rep/Group
		
		If IsEmpty(intRepID) Then
			SendEmail lngCaseID, "CASE_UPDATED_REP", "GROUP"
		Else
			SendEmail lngCaseID, "CASE_UPDATED_REP", "REP"
		End If
		
	Else
	
		' Do nothing
		
	End If	
	
Else

	' Do nothing, as EMail notification not enabled
	
End If
	

' Close case object

Set objCase = Nothing

		
' -------------------------------------------------------------------------------
' Display HTML
	
%>

<HTML>
<HEAD>
	
	<META content="MSHTML 6.00.2600.0" name=GENERATOR>
</HEAD>

<LINK rel="stylesheet" type="text/css" href="default.css">

<BODY>
<P align=center>
<TABLE class=Normal align=center cellSpacing=1 cellPadding=1 width="680" border=0 bgColor=white>
	<TR>
		<TD>
			<%
			Response.Write DisplayHeader
			%>
		</TD>
	</TR>
	<TR>
		<TD>
			<TABLE class="lhd_Box" cellSpacing=0 cellPadding=1 width="100%" border=0>
				<TR class="lhd_Heading1">
					<TD colspan=5><%=Lang("Case_Updated")%></TD>
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
					<TD colspan=3 align=left>
						<B><A href="caseModify.asp?ID=<%=lngCaseID%>"><%=Lang("Case")%> #<%=lngCaseID%></A></B>
						<BR><BR>
						Thank you.   Your updates have been submitted successfully and will be reviewed as soon as possible.
					</TD>
					<TD></TD>
				</TR>
				<TR>
					<TD></TD>
					<TD></TD>
					<TD></TD>
					<TD></TD>
					<TD></TD>
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
