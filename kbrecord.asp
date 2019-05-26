<%@Language="VBScript"%>
<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: kbRecord.asp
'  Date:     $Date: 2004/03/11 06:08:42 $
'  Version:  $Revision: 1.4 $
'  Purpose:  Allows for creation and/or modification of a knowledgebase record
' ----------------------------------------------------------------------------------
%>

<% 

Option Explicit

%>
<HTML>

<!--METADATA TYPE="TypeLib" NAME="Microsoft ActiveX Data Objects 2.5 Library" UUID="{00000205-0000-0010-8000-00AA006D2EA4}" VERSION="2.5"-->
<!--METADATA TYPE="TypeLib" NAME="Microsoft Scripting Runtime" UUID="{420B2830-E718-11CF-893D-00A0C9054228}" VERSION="1.0"-->
<!--METADATA TYPE="TypeLib" NAME="Microsoft CDO for Windows 2000 Library" UUID="{CD000000-8B95-11D1-82DB-00C04FB1625D}" VERSION="1.0"-->

<!-- #Include File = "Include/Settings.asp" -->
<!-- #Include File = "Include/Public.asp" -->

<!-- #Include File = "Classes/clsContact.asp" -->
<!-- #Include File = "Classes/clsKnowledgebase.asp" -->

<%

' Declare variables

Dim cnnDB
Dim intUserID, intLastUpdateByID, intEnteredByID
Dim binUserPermMask
Dim dteLastUpdate, dteEnteredDate
Dim lngKnowledgebaseID
Dim blnSave, blnIsActive
Dim strIssue, strCause, strResolution, strHeading, strEnteredBy, strLastUpdateBy
Dim objKnowledgebase, objContact


' Create connection to the database
Set cnnDB = CreateConnection

' Check Session variables
intUserID = GetUserID
binUserPermMask = GetUserPermMask

' Check if the user has rights to modify the knowledgebase
If (PERM_KB_MODIFY = (PERM_KB_MODIFY And binUserPermMask)) Or (PERM_KB_CREATE = (PERM_KB_CREATE And binUserPermMask)) Then

Else
	DisplayError 4, ""
End If


' Determine whether we are creating/update a record or not
If Request.Form("tbxSave") = "1" Then
	blnSave = True
Else
	blnSave = False
End If


If blnSave = True Then
  ' Update a Knowledgebase record

  lngKnowledgebaseID = CInt(Request.Form("tbxKnowledgebaseID"))
	strIssue = Request.Form("txtIssue")
	strCause = Request.Form("txtCause")
	strResolution = Request.Form("txtResolution")
	intEnteredByID = Request.Form("tbxEnteredByID")

	If Request.Form("chkIsActive") = "on" Then
		blnIsActive = lhd_True
	Else
		blnIsActive = lhd_False
	End If

'	dteEnteredDate = Request.Form("txtEnteredDate")
	dteLastUpdate = CDate(Now())
	intLastUpdateByID = intUserID
	
	
	' Now Save/Update the Knowledgebase record
	Set objKnowledgebase = New clsKnowledgebase

	objKnowledgebase.ID = lngKnowledgebaseID

	If Not objKnowledgebase.Load Then
		' Raise error, Knowledgebase record doesn't exist

	Else
		' Knowledgebase record exists
		
	End If

	objKnowledgebase.Issue = strIssue
	objKnowledgebase.Cause = strCause
	objKnowledgebase.Resolution = strResolution
	objKnowledgebase.EnteredByID = intEnteredByID
'	objKnowledgebase.EnteredDate =
	objKnowledgebase.IsActive = blnIsActive
	objKnowledgebase.LastUpdate = dteLastUpdate
	objKnowledgebase.LastUpdateByID = intLastUpdateByID

	If Not objKnowledgebase.Update Then
		' Failed to create/save user
							
	Else
		lngKnowledgebaseID = CInt(objKnowledgebase.ID)
						
    Response.Redirect "kbView.asp?ID=" & CStr(objKnowledgebase.ID)

	End If
	
	Set objKnowledgebase = Nothing


Else
  ' Create/Display a Knowledgebase record

	If Request.QueryString.Count = 0 Then
	
    ' Check if the user has rights to create a knowledgebase record
    If PERM_KB_CREATE = (PERM_KB_CREATE And binUserPermMask) Then

    Else
    	DisplayError 4, ""
    End If
    
		strHeading = Lang("New_Knowledgebase_Record")

		' Create a new record.  Populate the default values

		lngKnowledgebaseID = 0
		strIssue = ""
		strCause = ""
		strResolution = ""
		intEnteredByID = intUserID
'		dteEnteredDate =
		blnIsActive = lhd_True
		dteLastUpdate = CDate(Now())
		intLastUpdateByID = intUserID
			
		Set objContact = New clsContact
		
		objContact.ID = intEnteredByID

		If Not(objContact.Load) Then
      ' Raise error, Object failed to load
		Else
      strEnteredBy = objContact.UserName
		End If
		
		Set objContact = Nothing
			
	Else
	
		strHeading = "Modify Knowledgebase Record"

		' Edit a Knowledgebase record, i.e. Load record
	
		lngKnowledgebaseID = CInt(Request.QueryString("ID"))

    Set objKnowledgebase = New clsKnowledgebase
    
    objKnowledgebase.ID = lngKnowledgebaseID
    
    If Not(objKnowledgebase.Load) Then
      ' Raise error, Object failed to load
      
    Else
		  strIssue = objKnowledgebase.Issue
		  strCause = objKnowledgebase.Cause
		  strResolution = objKnowledgebase.Resolution
		  strEnteredBy = objKnowledgebase.EnteredBy.UserName
		  intEnteredByID = objKnowledgebase.EnteredByID
		  dteEnteredDate = objKnowledgebase.EnteredDate
		  blnIsActive = objKnowledgebase.IsActive
		  dteLastUpdate = objKnowledgebase.LastUpdate
		  strLastUpdateBy = objKnowledgebase.LastUpdateBy.UserName
		  intLastUpdateByID = objKnowledgebase.LastUpdateByID

		End If
		
    Set objKnowledgebase = Nothing

	End If

End If
	
%>

<HEAD>
	
	<META content="MstrHTML 6.00.2600.0" name=GENERATOR>
</HEAD>

<LINK rel="stylesheet" type="text/css" href="Default.css">

<BODY>
<P align=center>
<TABLE class=Normal>
	<TR>
		<TD>
		<%
		Response.Write DisplayHeader
		%>
		</TD>
	</TR>
	<TR>
		<TD>
			<TABLE class="Normal" cellSpacing=0>
			  <FORM name="frmKnowledgebase" action="kbRecord.asp" method="post">
    	  <INPUT id="tbxSave" name="tbxSave" type="hidden" value="1">
    	  <INPUT id="tbxKnowledgebaseID" name="tbxKnowledgebaseID" type="hidden" value="<%=lngKnowledgebaseID%>">
    	  <INPUT id="tbxEnteredByID" name="tbxEnteredByID" type="hidden" value="<%=intEnteredByID%>">
    	  <INPUT id="tbxLastUpdateByID" name="tbxLastUpdateByID" type="hidden" value="<%=intLastUpdateByID%>">
				<TR class="lhd_Heading1">
					<TD colspan=5 align=middle><%=strHeading%></TD>
				</TR>
				<TR>
					<TD width="20%"></TD>
					<TD width="25%"></TD>
					<TD width="10%"></TD>
					<TD width="20%"></TD>
					<TD width="25%"></TD>
				</TR>
				<TR>
					<TD><%=Lang("Reference_ID")%>:</TD>
					<TD>
					  <%
					  If lngKnowledgebaseID = 0 Then
					    Response.Write "( <I>auto generate...</I> )"
					  Else
					    Response.Write "KB" & Right("00000000" & CStr(lngKnowledgebaseID), 8)
					  End If
					  %>
					</TD>
					<TD></TD>
					<TD><%=Lang("Last_Updated")%>:</TD>
					<TD><%If lngKnowledgebaseID = 0 Then Response.Write "" Else Response.Write DisplayDateTime(dteLastUpdate) End If%></TD>
				</TR>
				<TR>
					<TD></TD>
					<TD></TD>
					<TD></TD>
					<TD><%=Lang("Last_Updated_By")%>:</TD>
					<TD><%If lngKnowledgebaseID = 0 Then Response.Write "" Else Response.Write strLastUpdateBy End If%></TD>
				</TR>
				<TR>
					<TD></TD>
					<TD></TD>
					<TD></TD>
					<TD></TD>
					<TD></TD>
				</TR>
				<TR>
					<TD valign="Top" style="PADDING-TOP: 4px;"><%=Lang("Issue")%>:</TD>
					<TD colspan="4"><TEXTAREA name="txtIssue" id="txtIssue" style="SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-TRACK-COLOR: white; WIDTH: 100%; HEIGHT: 140px"><%=strIssue%></TEXTAREA></TD>
				</TR>
				<TR>
					<TD valign="Top" style="PADDING-TOP: 4px;"><%=Lang("Cause")%>:</TD>
					<TD colspan="4"><TEXTAREA name=txtCause id=txtCause style="SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-TRACK-COLOR: white; WIDTH: 100%; HEIGHT: 105px"><%=strCause%></TEXTAREA></TD>
				</TR>
				<TR>
					<TD valign="Top" style="PADDING-TOP: 4px;"><%=Lang("Resolution")%>:</TD>
					<TD colspan="4"><TEXTAREA name=txtResolution id=txtResolution style="SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-TRACK-COLOR: white; WIDTH: 100%; HEIGHT: 140px"><%=strResolution%></TEXTAREA></TD>
				</TR>
				<TR>
					<TD></TD>
					<TD>
					  <%
					  If blnIsActive = False Then
					  %> 
	  				  <INPUT type="checkbox" id="chkIsActive" name="chkIsActive">&nbsp;Active
					  <%
					  Else
					  %> 
  					  <INPUT type="checkbox" id="chkIsActive" name="chkIsActive" checked>&nbsp;Active
  				  <%
					  End If
					  %>
					</TD>
					<TD colspan="3" align="Right"></TD>
					<TD></TD>
				</TR>
				<TR>
					<TD></TD>
					<TD style="FONT-SIZE: 8pt;">
            <BR>
            <%=Lang("Entered_By")%>:&nbsp;<%If lngKnowledgebaseID = 0 Then Response.Write "" Else Response.Write strEnteredBy End If%>
            <BR>
            <%=Lang("Entered_Date")%>:&nbsp;<%If lngKnowledgebaseID = 0 Then Response.Write "" Else Response.Write DisplayDateTime(dteEnteredDate) End If%>
          </TD>
					<TD></TD>
					<TD></TD>
					<TD align="Right" valign="Bottom"><INPUT style="WIDTH: 80px; BACKGROUND-COLOR: white" id=btnSave name=btnSave type=submit value="<%=Lang("Save")%>"></TD>
				</TR>
      	</FORM>
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
