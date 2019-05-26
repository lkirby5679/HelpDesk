<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: admEmailMsg.asp
'  Date:     $Date: 2004/03/10 08:21:12 $
'  Version:  $Revision: 1.5 $
'  Purpose:  Administration page for creating/modifing Email Messages
' ----------------------------------------------------------------------------------

%>
<% Option Explicit
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



<HTML>
<HEAD>

<META content="MSHTML 6.00.2600.0" name=GENERATOR></HEAD>

<LINK rel="stylesheet" type="text/css" href="Default.css">
<%
	Dim cnnDB
	Dim binUserPermMask, binRequiredPerm
	Dim blnSave, blnIsActive
  Dim objEMailMsg, objCollection
  Dim strEMailMsgType, strSubject, strBody
  Dim strMode, strIsActiveHTML, strHeading
  Dim intUserID, intLastUpdateByID, intEMailMsgID, intLangID
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

		intEMailMsgID = Cint(Request.Form("tbxEMailMsgID"))
		strSubject = Request.Form("tbxSubject")
		strBody = Request.Form("txtBody")
		intLangID = Cint(Request.Form("cbxLanguage"))

		If Request.Form("chkIsActive") = "on" Then
			blnIsActive = lhd_True
		Else
			blnIsActive = lhd_False
		End If

		dteLastUpdate = SQLDateFormat(Day(Now()), Month(Now()), Year(Now())) & " " & FormatDateTime(Now(), vbShortTime)
		intLastUpdateByID = intUserID

		
		' Check for required fields
		


		' Now save/update the Email Message details

		Set objEMailMsg = New clsEMailMsg

		objEMailMsg.ID = intEMailMsgID

		If Not objEMailMsg.Load Then
		
			' Email Message does not exist

		Else
		
			' Email Message exists
		
		End If

		' Check the the fields and leave Null if nothing is set.

		objEMailMsg.LangID = intLangID
		objEMailMsg.Subject = strSubject
		objEMailMsg.Body = strBody
		objEMailMsg.IsActive = blnIsActive
		objEMailMsg.LastUpdate = dteLastUpdate
		objEMailMsg.LastUpdateByID = intLastUpdateByID

						
		If Not objEMailMsg.Update Then
						
			' Failed to create/save Email Message

							
		Else
						
			intEMailMsgID = objEMailMsg.ID
			strHeading = Lang("Email_Message_Saved")
						
		End If
						
		Set objEMailMsg = Nothing
%>
		<TR>
		   <TD>
		      <TABLE class=Normal cellSpacing=0 cellPadding=1 width="100%" border=0 bgColor=white>
		         <TR class="lhd_Heading1">
					<TD></TD>
 					<TD colspan=3 align=middle><%=strHeading%></TD>
					<TD></TD>
				 </TR>
		      <TR>
		        <TD align=right colspan=5><a style="FONT-SIZE: 8pt" href="admEmailMsgList.asp?Page=1"><%=Lang("Manage_Email_Messages")%></a></TD>
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
 					<TD colspan=3 align=left>Email Message information has been successfully saved.</TD>
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
	
		' Mode: 1 - To create a new Email Message
		'		2 - Edit a Email Message
	
		If Request.QueryString.Count = 0 Then
	
			' Create a new record
	
			strMode = 1
			intEMailMsgID = 0
			strHeading = Lang("New_Email_Message")
			
		Else
	
			' Edit a record determine by the Email Message ID passed via the QueryString
	
			strMode = 2
			intEMailMsgID = Request.QueryString("id")
			strHeading = Lang("Modify_Email_Message")
			
		End If
			

		Select Case strMode
		
			Case 1	' Create new EMailMessage.  We should be able to do this.
			
				strSubject = ""
				strBody = ""
				intLangID = Application("DEFAULT_LANGUAGE")

				blnIsActive = lhd_True
				strIsActiveHTML = "CHECKED"

				dteLastUpdate = ""
				intLastUpdateByID = 0


			Case 2  ' Edit Email Message

				' Get the Email Message t ID we want to edit and load the record

				Set objEMailMsg = New clsEMailMsg

				objEMailMsg.ID = intEMailMsgID
			
				If Not objEMailMsg.Load Then
				
					' Couldn't load user for some reason
					Response.Write objEMailMsg.LastError
				
				Else
				
					strEMailMsgType = objEMailMsg.EMailMsgType
					intLangID = objEMailMsg.LangID
					strSubject = objEMailMsg.Subject
					strBody = objEMailMsg.Body
					
					If objEMailMsg.IsActive = True Then
						strIsActiveHTML = "CHECKED"
					Else
						strIsActiveHTML = ""
					End If
					
					dteLastUpdate = objEMailMsg.LastUpdate
					intLastUpdateByID = objEMailMsg.LastUpdateByID

				End If

				Set objEMailMsg = Nothing
				

			Case Else
				' Do nothing
				
		End Select

%>

  <TR>
    <TD>
	  <TABLE class=Normal cellSpacing=0 cellPadding=1 width="100%" border=0 bgColor=white>
	  <FORM action="admEmailMsg.asp" method="post" id=frmEMailMsg name=frmEMailMsg>
	  <INPUT id=tbxEMailMsgID name=tbxEMailMsgID type=hidden value="<%=intEMailMsgID%>">
	  <INPUT id=tbxSave name=tbxSave type=hidden value="1">
	  
		<TR class="lhd_Heading1">
			<TD colspan=5 align=center><%=strHeading%></TD>
		</TR>
		<TR>
		  <TD align=right colspan=5><a style="FONT-SIZE: 8pt" href="admEmailMsgList.asp?Page=1"><%=Lang("Manage_Email_Messages")%></a></TD>
		</TR>
		<TR>
		  <TD width="22%"><%=Lang("Email_Message_Type")%>:</TD>
		  <TD width="25%"><%=strEMailMsgType%></TD>
		  <TD width="5%"></TD>
		  <TD width="18%"></TD>
		  <TD width="25%"></TD>
		</TR>
		<TR>
		  <TD class="lhd_Required"><%=Lang("Subject")%>:</TD>
		  <TD colspan=4><INPUT id=tbxSubject name=tbxSubject style="WIDTH: 100%" value="<%=strSubject%>"></TD>
		</TR>
		<TR>
		  <TD class="lhd_Required" valign=top><%=Lang("Body")%>:</TD>
		  <TD colspan=4><TEXTAREA id=txtBody style="SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-TRACK-COLOR: white; WIDTH: 100%; HEIGHT: 140px" name=txtBody><%=strBody%></TEXTAREA></TD>
		</TR>
		<TR>
			<TD></TD>
			<TD></TD>
			<TD></TD>
			<TD class="lhd_Required"><%=Lang("Language")%>:</TD>
			<TD>
			     <SELECT id=cbxLanguage name=cbxLanguage style="WIDTH: 100%">
					<%
					Set objCollection = New clsCollection
			   
					objCollection.CollectionType = objCollection.clLanguage
					objCollection.Query = "SELECT LangPK, LangName FROM tblLanguages WHERE IsActive=" & lhd_True & " ORDER BY LangName ASC"
												
					If Not objCollection.Load Then
				
						Response.Write objCollection.LastError
							
					Else
				
					    Do While Not objCollection.EOF
						    
							If objCollection.Item.ID =  CInt(intLangID) Then
							%>
								<OPTION SELECTED VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.LangName%></OPTION>
							<%
							Else
							%>
								<OPTION VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.LangName%></OPTION>
							<%
							End If
								
							objCollection.MoveNext
								
					    Loop
						    
					End If
											
					Set objCollection = Nothing
					%>
			     </SELECT>
			</TD>
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
  </P>
  </BODY>
  </HTML>
  
<%
  
cnnDB.Close
Set cnnDB = Nothing
  
%>
