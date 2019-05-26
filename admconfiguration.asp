<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: admConfiguration.asp
'  Date:     $Date: 2004/03/30 04:20:30 $
'  Version:  $Revision: 1.9 $
'  Purpose:  Administration page for setting the system configuration settings
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
	Dim binUserPermMask
	Dim blnSave
'	Dim blnIsActive
  Dim objParam, objCollection
'  Dim strParamName, strParamValue, strMode, strIsActiveHTML
  Dim intUserID
'  Dim intLastUpdateByID, intParamID
'  Dim dteLastUpdate
'  Dim intDefaultRoleID, intDefaultStatusID, intDefaultPriorityID, intDefaultContactTypeID
'  Dim intDefaultLanguageID, intStatusClosed, intStatusOpen, intAttachmentMethod
'  Dim strSiteName
    

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

		'  For the saving to work we will need to maintain the Field names to be exactly
		' the same as those ParamNames stored in the tblParameters. 

		Set objParam = New clsParameter
		
		objParam.SetValue "AUTH_TYPE", CInt(Request.Form("cbxAuthenticationType"))
		objParam.SetValue "DATE_FORMAT", Trim(Request.Form("cbxDateFormat"))
		objParam.SetValue "DEFAULT_LANGUAGE", CInt(Request.Form("cbxDefaultLanguage"))
		objParam.SetValue "DEFAULT_PRIORITY", CInt(Request.Form("cbxDefaultPriority"))
		objParam.SetValue "DEFAULT_ROLE", CInt(Request.Form("cbxDefaultRole"))
		objParam.SetValue "DEFAULT_STATUS", CInt(Request.Form("cbxDefaultStatus"))
		objParam.SetValue "ENABLE_ATTACHMENTS", CInt(Request.Form("cbxEnableAttachments"))
		objParam.SetValue "ENABLE_INOUT", CInt(Request.Form("cbxEnableInOut"))
		objParam.SetValue "ENABLE_KB", CInt(Request.Form("cbxEnableKB"))
		objParam.SetValue "ENABLE_EMAIL", CInt(Request.Form("cbxEnableEMail"))
		objParam.SetValue "ENABLE_REPORTS", CInt(Request.Form("cbxEnableReports"))
		objParam.SetValue "EMAIL_METHOD", CInt(Request.Form("cbxEMailMethod"))
		objParam.SetValue "ITEMS_PER_PAGE", CInt(Request.Form("tbxItemsPerPage"))
		objParam.SetValue "MAX_ATTACHMENT_SIZE", CLng(Request.Form("tbxMaxAttachmentSize"))
		objParam.SetValue "SITE_NAME", Trim(Request.Form("tbxSiteName"))
		objParam.SetValue "SMTP_SERVER", Trim(Request.Form("tbxSMTPServer"))
		objParam.SetValue "SYSTEM_EMAIL", Trim(Request.Form("tbxSystemEMail"))
		objParam.SetValue "STATUS_CLOSED", CInt(Request.Form("cbxStatusClosed"))
		objParam.SetValue "STATUS_OPEN", CInt(Request.Form("cbxStatusOpen"))
		objParam.SetValue "STATUS_CANCELLED", CInt(Request.Form("cbxStatusCancelled"))
		objParam.SetValue "TIME_FORMAT", Trim(Request.Form("cbxTimeFormat"))
		
		Set objParam = Nothing

	%>
		<TR>
		   <TD>
		      <TABLE class=Normal width="100%" border=0 cellSpacing=0 cellPadding=1>
				    <TR class="lhd_Heading1">
				       <TD colspan=5 align=center><%=Lang("System_Configuration_Saved")%></TD>
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
 				       <TD colspan=3 align=Left>The System Configuration has successfully been saved.</TD>
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
	
		Set objParam = New clsParameter

	%>

  <TR>
    <TD>
	  <TABLE class=Normal cellSpacing=0 cellPadding=1 width="100%" border=0 bgColor=white>
	  <FORM action="admConfiguration.asp" method="post" id=frmConfiguration name=frmConfiguration>
	  <INPUT id=tbxSave name=tbxSave type=hidden value="1">
	  
		<TR class="lhd_Heading1">
			<TD colspan=5 align=center><%=Lang("System_Configuration")%></TD>
		</TR>
		<TR>
		   <TD colspan=5></TD>
		</TR>
		<TR class="lhd_Heading2">
			<TD colspan="5"><%=Lang("Site_Information")%></TD>
		</TR>
		<TR>
		  <TD width="25%"><%=Lang("Site_Name")%>:</TD>
		  <TD width="24%"><INPUT id=tbxSiteName style="WIDTH: 100%" name=tbxSiteName value="<%=objParam.GetValue("SITE_NAME")%>"></TD>
		  <TD width="5%"></TD>
		  <TD width="25%"><%=Lang("Version")%>:</TD>
		  <TD width="21%"><%=objParam.GetValue("VERSION")%></TD>
		</TR>
		<TR>
		  <TD><%=Lang("System_Email_Address")%>:</TD>
		  <TD><INPUT id=tbxSystemEmail style="WIDTH: 100%" name=tbxSystemEmail value="<%=objParam.GetValue("SYSTEM_EMAIL")%>"></TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		</TR>
		<TR>
		  <TD><%=Lang("Date_Format")%>:</TD>
      <TD>
			  <SELECT id=cbxDateFormat style="WIDTH: 100%" name=cbxDateFormat>
			    <%
			    If objParam.GetValue("DATE_FORMAT") = "dd/mm/yyyy" Then
			    %>
			  	  <OPTION value="dd/mm/yyyy" selected>dd/mm/yyyy</OPTION>
			  	  <OPTION value="mm/dd/yyyy">mm/dd/yyyy</OPTION>
			  	  <OPTION value="dd-mmm-yyyy">dd-mmm-yyyy</OPTION>
			  	  <OPTION value="dd.mm.yyyy">dd.mm.yyyy</OPTION>
			  	  <OPTION value="mm.dd.yyyy">mm.dd.yyyy</OPTION>
			  	<%
			  	ElseIf objParam.GetValue("DATE_FORMAT") = "mm/dd/yyyy" Then
			  	%>
			  	  <OPTION value="dd/mm/yyyy">dd/mm/yyyy</OPTION>
			  	  <OPTION value="mm/dd/yyyy" selected>mm/dd/yyyy</OPTION>
			  	  <OPTION value="dd-mmm-yyyy">dd-mmm-yyyy</OPTION>
			  	  <OPTION value="dd.mm.yyyy">dd.mm.yyyy</OPTION>
			  	  <OPTION value="mm.dd.yyyy">mm.dd.yyyy</OPTION>
			  	<%
			  	ElseIf objParam.GetValue("DATE_FORMAT") = "dd-mmm-yyyy" Then
			  	%>
			  	  <OPTION value="dd/mm/yyyy">dd/mm/yyyy</OPTION>
			  	  <OPTION value="mm/dd/yyyy">mm/dd/yyyy</OPTION>
			  	  <OPTION value="dd-mmm-yyyy" selected>dd-mmm-yyyy</OPTION>
			  	  <OPTION value="dd.mm.yyyy">dd.mm.yyyy</OPTION>
			  	  <OPTION value="mm.dd.yyyy">mm.dd.yyyy</OPTION>
			  	<%
			  	ElseIf objParam.GetValue("DATE_FORMAT") = "dd.mm.yyyy" Then
			  	%>
			  	  <OPTION value="dd/mm/yyyy">dd/mm/yyyy</OPTION>
			  	  <OPTION value="mm/dd/yyyy">mm/dd/yyyy</OPTION>
			  	  <OPTION value="dd-mmm-yyyy">dd-mmm-yyyy</OPTION>
			  	  <OPTION value="dd.mm.yyyy" selected>dd.mm.yyyy</OPTION>
			  	  <OPTION value="mm.dd.yyyy">mm.dd.yyyy</OPTION>
			  	<%
			  	Else
			  	%>
			  	  <OPTION value="dd/mm/yyyy">dd/mm/yyyy</OPTION>
			  	  <OPTION value="mm/dd/yyyy">mm/dd/yyyy</OPTION>
			  	  <OPTION value="dd-mmm-yyyy">dd-mmm-yyyy</OPTION>
			  	  <OPTION value="dd.mm.yyyy">dd.mm.yyyy</OPTION>
			  	  <OPTION value="mm.dd.yyyy" selected>mm.dd.yyyy</OPTION>
	  	    <%
			  	End If
			  	%>
			  </SELECT>
		  </TD>
		  <TD></TD>
		  <TD><%=Lang("Time_Format")%>:</TD>
      <TD>
			  <SELECT id=cbxTimeFormat style="WIDTH: 100%" name=cbxTimeFormat>
			  	<OPTION value="hh:mm">hh:mm</OPTION>
			  </SELECT>
		  </TD>
		</TR>
		<TR>
		  <TD><%=Lang("Items_Per_Page")%>:</TD>
		  <TD><INPUT id="tbxItemsPerPage" style="WIDTH: 100%" name="tbxItemsPerPage" value="<%=objParam.GetValue("ITEMS_PER_PAGE")%>"></TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
    </tr>
		<TR>
		  <TD><%=Lang("Attachments_Enabled")%>:</TD>
      <TD>
			  <SELECT id="cbxEnableAttachments" style="WIDTH: 100%" name=cbxEnableAttachments>
			  	<%
			  	If  CInt(objParam.GetValue("ENABLE_ATTACHMENTS")) = 1 Then
			  	%>
			  	<OPTION selected value="1">Yes</OPTION>
			  	<OPTION value="0">No</OPTION>
			  	<%
			  	Else
			  	%>
			  	<OPTION value="1">Yes</OPTION>
			  	<OPTION selected value="0">No</OPTION>
			  	<%
			  	End If
			  	%>
			  </SELECT>
		  </TD>
		  <TD></TD>
		  <TD><%=Lang("Max_Attachment_Size")%>:</TD>
		  <TD><INPUT id="tbxMaxAttachmentSize" style="WIDTH: 100%" name=tbxMaxAttachmentSize value="<%=objParam.GetValue("MAX_ATTACHMENT_SIZE")%>"></TD>
    </tr>
		<TR>
			<TD colspan=5></TD>
		</TR>
		<TR class="lhd_Heading2">
			<TD colspan="5"><%=Lang("Authentication")%></TD>
		</TR>
		<TR>
		  <TD><%=Lang("Authentication_Type")%>:</TD>
      <TD>
			  <INPUT type=hidden id=tbxAuthenticationType name=tbxAuthenticationType value="<%=objParam.GetValue("AUTH_TYPE")%>">
			  <SELECT id=cbxAuthenticationType style="WIDTH: 100%" name=cbxAuthenticationType>
			  	<OPTION value="1" selected>NT/AD Authentication</OPTION>
			  	<OPTION value="2">DB (Huron) Authentication</OPTION>
			  </SELECT>
			  <SCRIPT language="VBScript">
			  	document.frmConfiguration.cbxAuthenticationType.selectedIndex = CInt(document.frmConfiguration.tbxAuthenticationType.value) - 1
			  </SCRIPT>
		  </TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		</TR>
		<TR>
			<TD colspan=5></TD>
		</TR>
		<TR class="lhd_Heading2">
			<TD colspan="5"><%=Lang("Add-Ons")%></TD>
		</TR>
		<TR>
		  <TD><%=Lang("Knowledgebase_Enabled")%>:</TD>
      <TD>
			  <SELECT id=cbxEnableKB style="WIDTH: 100%" name=cbxEnableKB>
			  	<%
			  	If  CInt(objParam.GetValue("ENABLE_KB")) = 1 Then
			  	%>
			  	<OPTION selected value="1">Yes</OPTION>
			  	<OPTION value="0">No</OPTION>
			  	<%
			  	Else
			  	%>
			  	<OPTION value="1">Yes</OPTION>
			  	<OPTION selected value="0">No</OPTION>
			  	<%
			  	End If
			  	%>
			  </SELECT>
		  </TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		</TR>
		<TR>
		  <TD><%=Lang("Reports_Enabled")%>:</TD>
      <TD>
			  <SELECT id=cbxEnableReports style="WIDTH: 100%" name=cbxEnableReports>
			  	<%
			  	If CInt(objParam.GetValue("ENABLE_REPORTS")) = 1 Then
			  	%>
			  	<OPTION value="1" selected>Yes</OPTION>
			  	<OPTION value="0">No</OPTION>
			  	<%
			  	Else
			  	%>
			  	<OPTION value="1">Yes</OPTION>
			  	<OPTION value="0" selected>No</OPTION>
			  	<%
			  	End If
			  	%>
			  </SELECT>
		  </TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		</TR>
		<TR>
		  <TD><%=Lang("In/Out_Phone_List_Enabled")%>:</TD>
      <TD>
			  <SELECT id=cbxEnableInOut style="WIDTH: 100%" name=cbxEnableInOut>
			  	<%
			  	If  CInt(objParam.GetValue("ENABLE_INOUT")) = 1 Then
			  	%>
			  	<OPTION selected value="1">Yes</OPTION>
			  	<OPTION value="0">No</OPTION>
			  	<%
			  	Else
			  	%>
			  	<OPTION value="1">Yes</OPTION>
			  	<OPTION selected value="0">No</OPTION>
			  	<%
			  	End If
			  	%>
			  </SELECT>
		  </TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		</TR>
		<TR>
			<TD colspan=5></TD>
		</TR>
		<TR class="lhd_Heading2">
			<TD colspan="5"><%=Lang("Notification_Settings")%></TD>
		</TR>
		<TR>
		  <TD><%=Lang("Email_Enabled")%>:</TD>
		  <TD>
				<SELECT id=cbxEnableEMail style="WIDTH: 100%" name=cbxEnableEMail>
					<%
					If CInt(objParam.GetValue("ENABLE_EMAIL")) = 1 Then
					%>
					<OPTION value="1" selected>Yes</OPTION>
					<OPTION value="0">No</OPTION>
					<%
					Else
					%>
					<OPTION value="1">Yes</OPTION>
					<OPTION value="0" selected>No</OPTION>
					<%
					End If
					%>
				</SELECT>
		  </TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		</TR>
		<TR>
		  <TD><%=Lang("Email_Type")%>:</TD>
      <TD>
			  <INPUT type=hidden id=tbxEMailMethod name=tbxEMailMethod value="<%=objParam.GetValue("EMAIL_METHOD")%>">
			  <SELECT id=cbxEMailMethod style="WIDTH: 100%" name=cbxEMailMethod>
			  	<OPTION value="1">CDOSYS</OPTION>
			  	<OPTION value="2">CDONTS</OPTION>
			  	<OPTION value="3">JMail</OPTION>
			  	<OPTION value="4">ASP Mail</OPTION>
			  </SELECT>
			  <SCRIPT language="VBScript">
			  	document.frmConfiguration.cbxEMailMethod.selectedIndex = CInt(document.frmConfiguration.tbxEMailMethod.value) - 1
			  </SCRIPT>
		  </TD>
		  <TD></TD>
		  <TD><%=Lang("SMTP_Server")%>:</TD>
		  <TD><INPUT id=tbxSMTPServer name=tbxSMTPServer style="WIDTH: 100%" value="<%=objParam.GetValue("SMTP_SERVER")%>"></TD>
		</TR>
		<TR>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		  <TD></TD>
		</TR>
		<TR class="lhd_Heading2">
			<TD colspan="5"><%=Lang("System_Defaults")%></TD>
		</TR>
		<TR>
		  <TD><%=Lang("Open_Status")%>:</TD>
      <TD>
			  <SELECT id=cbxStatusOpen style="WIDTH: 100%" name=cbxStatusOpen>
  			  <%
	  		  Response.Write BuildList("STATUS_LIST", CInt(objParam.GetValue("STATUS_OPEN")))
		      %>
			  </SELECT>
		  </TD>
		  <TD></TD>
		  <TD><%=Lang("Default_Priority")%>:</TD>
      <TD>
			  <SELECT id=cbxDefaultPriority style="WIDTH: 100%" name=cbxDefaultPriority>
			    <%
			    Response.Write BuildList("PRIORITY_LIST", CInt(objParam.GetValue("DEFAULT_PRIORITY")))
		      %>
			  </SELECT>
		  </TD>
		</TR>
		<TR>
		  <TD><%=Lang("Cancelled_Status")%>:</TD>
      <TD>
  			<SELECT id=cbxStatusCancelled style="WIDTH: 100%" name=cbxStatusCancelled>
	  		  <%
		  	  Response.Write BuildList("STATUS_LIST", CInt(objParam.GetValue("STATUS_CANCELLED")))
		      %>
  			</SELECT>
		  </TD>
		  <TD></TD>
		  <TD><%=Lang("Default_Status")%>:</TD>
      <TD>
			  <SELECT id=cbxDefaultStatus style="WIDTH: 100%" name=cbxDefaultStatus>
			    <%
			    Response.Write BuildList("STATUS_LIST", CInt(objParam.GetValue("DEFAULT_STATUS")))
		      %>
			  </SELECT>
		  </TD>
		</TR>
		<TR>
		  <TD><%=Lang("Closed_Status")%>:</TD>
      <TD>
			  <SELECT id=cbxStatusClosed style="WIDTH: 100%" name=cbxStatusClosed>
			    <%
			    Response.Write BuildList("STATUS_LIST", CInt(objParam.GetValue("STATUS_CLOSED")))
		      %>
			  </SELECT>
		  </TD>
		  <TD></TD>
		  <TD><%=Lang("Default_Role")%>:</TD>
  		<TD>
		    <SELECT id=cbxDefaultRole name=cbxDefaultRole style="WIDTH: 100%"> 
		       <%
		    	Set objCollection = New clsCollection
		       
		    	objCollection.CollectionType = objCollection.clRole
		    	objCollection.Query = "SELECT RolePK, RoleName FROM tblRoles WHERE IsActive=" & lhd_True & " ORDER BY RolePK ASC"
		    							
		    	If Not objCollection.Load Then
		    	
		    		Response.Write objCollection.LastError
		    		
		    	Else
		    	
		    	    Do While Not objCollection.EOF
		    	    
		    			If objCollection.Item.ID =  CInt(objParam.GetValue("DEFAULT_ROLE")) Then
		    			%>
		    				<OPTION SELECTED VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.RoleName%></OPTION>
		    			<%
		    			Else
		    			%>
		    				<OPTION VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.RoleName%></OPTION>
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
		  <TD></TD>
		  <TD></TD>
		  <TD><%=Lang("Default_Language")%>:</TD>
		  <TD>
		    <SELECT id=cbxDefaultLanguage name=cbxDefaultLanguage style="WIDTH: 100%"> 
		       <%
		    	Set objCollection = New clsCollection
		       
		    	objCollection.CollectionType = objCollection.clLanguage
		    	objCollection.Query = "SELECT LangPK, LangName FROM tblLanguages WHERE IsActive=" & lhd_True & " ORDER BY LangPK ASC"
		    							
		    	If Not objCollection.Load Then
		    	
		    		Response.Write objCollection.LastError
		    		
		    	Else
		    	
		    	    Do While Not objCollection.EOF
		    	    
		    			If objCollection.Item.ID =  CInt(objParam.GetValue("DEFAULT_LANGUAGE")) Then
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
		  <TD></TD>
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
	  </TABLE>
  <%
  
		Set objParam = Nothing
  
  End If	' End Save
  %>
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
  </BODY></HTML>
  
<%
  
cnnDB.Close
Set cnnDB = Nothing
  
%>
  
  
