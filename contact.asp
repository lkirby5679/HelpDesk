<% Option Explicit %>

<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: Contact.asp
'  Date:     $Date: 2004/03/11 06:08:42 $
'  Version:  $Revision: 1.5 $
'  Purpose:  Allows a contact to update their profile information as well as allowing
'            a new contact to register.
' ----------------------------------------------------------------------------------
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
  
  <META content="MSHTML 6.00.2600.0" name=GENERATOR>
</HEAD>

<LINK rel="stylesheet" type="text/css" href="default.css">

<%

Dim cnnDB
Dim blnSave
Dim objContact, objCollection
Dim dtIOStatusDate, dtCreated, dteLastUpdate, dtLastAccess
Dim binUserPermMask

Dim strFName, strLName, strOfficeLocation, strPagerEmail
Dim strJobTitle, strOfficePhone, strHomePhone, strMobilePhone
Dim strMode, strUserName, strIOStatusText, strResume, strEmail, strNotes
Dim strHeader, strPassword

Dim intUserID, intPhotoFileID, intTZOffset, intLastUpdateByID, intJobFunctionID
Dim intContactTypeID, intDeptID, intOrgID, intLangID, intContactID, intIOStatusID


' Create database connection
Set cnnDB = CreateConnection
   
%>	
<BODY>

<P align=center>

<TABLE align=center cellSpacing=1 cellPadding=1 width="680" border=0>
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

  intContactID = Request.Form("tbxContactID")
  strUserName = Request.Form("tbxUserName")
  strPassword = Request.Form("pwdPassword")
  strFName = Request.Form("tbxFName")
  strLName = Request.Form("tbxLName")
  intContactTypeID = CInt(Request.Form("cbxContactType"))
  intOrgID = CInt(Request.Form("cbxOrganisation"))
  intDeptID = Cint(Request.Form("cbxDepartment"))
  intLangID = CInt(Request.Form("cbxLanguage"))
  strHomePhone = Request.Form("tbxHomePhone")
  strOfficePhone = Request.Form("tbxOfficePhone")
  strMobilePhone = Request.Form("tbxMobilePhone")
  strJobTitle = Request.Form("tbxJobTitle")
  intJobFunctionID = CInt(Request.Form("cbxJobFunction"))
  strEmail = Request.Form("tbxEmail")
  strPagerEmail = Request.Form("tbxPagerEmail")
  strOfficeLocation = Request.Form("tbxOfficeLocation")
  strNotes = Request.Form("txtNotes")

  intIOStatusID = CInt(Request.Form("cbxIOStatus"))
  dtIOStatusDate = SQLDate( Request.Form("tbxIOStatusDate") )
  strIOStatusText = Request.Form("tbxIOStatusText")
  intTZOffset = CInt(Request.Form("cbxTZOffset"))
	
  strResume  = Request.Form("txtResume")
  intPhotoFileID = CInt(Request.Form("cbxPhotoFile"))
'  dteLastUpdate = SQLDate( Day(Now) & "/" & Month(Now) & "/" & Year(Now) & " " & FormatDateTime(Now, vbShortTime) )
  dteLastUpdate = SQLDateFormat(Day(Now()), Month(Now()), Year(Now())) & " " & FormatDateTime(Now(), vbShortTime)
  intLastUpdateByID = intUserID
		
  ' Need to check for required fields
		
  If Len(strFName) = 0 Or Len(strLName) = 0 Or Len(strEmail) = 0 Or intDeptID = 0 Then
		
    Call DisplayError(1, "All required fields need to be entered, please go back and populate these fields")
    
  Else
  
    ' Do nothing
			
  End If


  ' Now save/update the contacts details

  Set objContact = New clsContact

  objContact.ID = intContactID

  If Not objContact.Load Then
		
    ' Contact does not exist

  Else
		
    ' Contact exists

  End If
		
  objContact.UserName = strUserName
  objContact.FName = strFName
  objContact.LName = strLName

  If intContactTypeID > 0 Then
    objContact.ContactTypeID  = intContactTypeID
  End If

  If intOrgID > 0 Then
    objContact.OrgID = intOrgID
  End If

  If intDeptID > 0 Then
    objContact.DeptID = intDeptID
  End If
		
  objContact.LangID = intLangID
  objContact.HomePhone = strHomePhone
  objContact.OfficePhone = strOfficePhone
  objContact.MobilePhone = strMobilePhone
  objContact.JobTitle = strJobTitle

  If intJobFunctionID > 0 Then
    objContact.JobFunctionID = intJobFunctionID
  End If

  objContact.Email = strEmail
  objContact.Password = strPassword

  objContact.PagerEmail = strPagerEmail
  objContact.OfficeLocation = strOfficeLocation
  objContact.Notes = strNotes

  If intIOStatusID > 0 Then
    objContact.IOStatusID = intIOStatusID
  End If

  If IsDate(dtIOStatusDate) Then
    objContact.IOStatusDate = dtIOStatusDate
  End If

  objContact.IOStatusText = strIOStatusText

  If intTZOffset > 0 Then
	objContact.TZOffset = intTZOffset
  End If

  If IsDate(dtLastAccess) Then
    objContact.LastAccess = dtLastAccess
  End If

  objContact.sResume = strResume

  If intPhotoFileID > 0 Then
    objContact.PhotoFileID = intPhotoFileID
  End If

  If IsDate(dteLastUpdate) Then
    objContact.LastUpdate = dteLastUpdate
  End If

  If intLastUpdateByID > 0 Then
    objContact.LastUpdateByID = intLastUpdateByID
  End If
					
  If Not objContact.Update Then
    ' Raise Error, Failed to create/save user
  Else
    intContactID = objContact.ID
  End If
						
  Set objContact = Nothing
  
%>
<TR>
   <TD>
      <TABLE class=Normal cellSpacing=1 cellPadding=1 width="100%" border=0 bgColor=white>
		 <TR class="lhd_Heading1">
			<TD colspan=5 align=center><%=Lang("Saved")%></TD>
		 </TR>
         <TR>
		 <TD colspan=5 align=center></TD>
		 </TR>
         <TR>
		 <TD colspan=5 align=center><A href="Contact.asp?id=<%=intContactID%>"><%=strUsername & " (" & strFName & " " & strLName & ")"%></A>&nbsp;<%=Lang("Details Saved")%></TD>
		 </TR>
         <TR>
		 <TD colspan=5 align=center></TD>
		 </TR>
      </TABLE>
   </TD>
</TR>
<%

Else
	
  If Request.QueryString.Count = 0 Then
	
    ' Create a new record
	
    strMode = 1
    intContactID = 0
    strHeader = Lang("Register")
    			
  Else
    	
    ' Edit a contact record
    	
    strMode = 2
    intContactID = Request.QueryString("id")
    strHeader = Lang("Modify") & " " & Lang("Contact")
    			
  End If
    		
  Select Case strMode
    		
    Case 1	' Register new Contact logging in
    			
      Select Case Application("AUTH_TYPE")
          					
        Case lhd_ADAuthentication	' AD or NT Authentication

          '  Connect through to LDAP/AD database on the windows domain and extract the required
          ' information

          Dim objADsRootDSE
          Dim objUser
          Dim cnnADOConnection
          Dim cmdADOCommand
          Dim strADsPath, strBase, strName, strObjects, strFilter, strAttributes, strScope
          Dim rstADORecordset

'
'         GC:// is specifically for AD
'
'          ' First, need to discover the local global catalog server
'          Set objADsRootDSE = GetObject("GC://RootDSE")
'
'          ' Form an ADsPath string to the DN of the root of the Active Directory forest
'          strADsPath = "GC://" & objADsRootDSE.Get("rootDomainNamingContext")
            
          ' First, need to discover the local global catalog server
          Set objADsRootDSE = GetObject("LDAP://RootDSE")

          ' Form an ADsPath string to the DN of the root of the Active Directory forest
          strADsPath = "LDAP://" & objADsRootDSE.Get("rootDomainNamingContext")

          ' Wrap the ADsPath with angle brackets to form the base string
          strBase = "<" & strADsPath & ">"
          							
          ' Release the ADSI object, no longer needed
          Set objADsRootDSE = Nothing
          							
          '  Specify the LDAP filter First, indicate the category of objects to
          ' be searched (all people, not just users)
          strObjects = "(objectCategory=person)"

          ' Strip the domain part
          strName = Right(Request.ServerVariables("AUTH_USER"), Len(Request.ServerVariables("AUTH_USER")) - InStr(Request.ServerVariables("AUTH_USER"), "\"))

          ' Add the two filters together
          strFilter = "(&" & strObjects & "sAMAccountName=" & strName & ")"

          '  Set the attributes we want the recordset to contain.  We're interested in
          ' the common name and telephone number
          strAttributes = "cn, adspath"

          ' Specify the scope (base, onelevel, subtree)
          strScope = "subtree"

          ' Create ADO connection using the ADSI OLE DB provider
          Set cnnADOConnection = Server.CreateObject("ADODB.Connection")
          cnnADOConnection.Open "Provider=ADsDSOObject"

          ' Create ADO commmand object and associate it with the connection
          Set cmdADOCommand = Server.CreateObject("ADODB.Command")
          cmdADOCommand.ActiveConnection = cnnADOConnection

          ' Create the command string using the four parts
          cmdADOCommand.CommandText = strBase & ";" & strFilter & ";" & strAttributes & ";" & strScope

          ' Execute the query for the user in the directory
          Set rstADORecordset = cmdADOCommand.Execute

          If rstADORecordset.BOF And rstADORecordset.EOF Then
          							
            Response.Write "No records were found."
          								
          Else
          							
            ' Here we are only going to use the first record we find and disregard the
            ' rest.  This is really only for auto-filling fields and there really should
            ' be only one user per GC
            							
            ' Create the user object to get the properties
            Set objUser = GetObject(rstADORecordset.Fields("adspath"))

            strUserName = strName
            strPassword = ""
            strFName = objUser.givenName
            strLName = objUser.sn
            intContactTypeID = Application("DEFAULT_CONTACT_TYPE")
            intOrgID = 0
            intDeptID = 0
            intLangID = Application("DEFAULT_LANGUAGE")
            strHomePhone = objUser.homePhone
            strOfficePhone = objUser.telephoneNumber
            strMobilePhone = objUser.mobile
            strJobTitle = objUser.title
            intJobFunctionID = 0
            strEmail = objUser.mail
            strPagerEmail = ""
            strOfficeLocation = objUser.physicalDeliveryOfficeName
            strNotes = ""
            intIOStatusID = 0
            dtIOStatusDate = ""
            strIOStatusText = ""
            intTZOffset = 0
            strResume  = ""
            intPhotoFileID = 0
            dtLastAccess = ""
            dteLastUpdate = ""
            intLastUpdateByID = 0
            								
            Set objUser = Nothing
          								
          End If

          rstADORecordset.Close
          Set rstADORecordset = Nothing

          Set cmdADOCommand = Nothing

          cnnADOConnection.Close
          Set cnnADOConnection = Nothing
          
    							
        Case lhd_DBAuthentication	' DB Authentication

          ' Set default fields

          strUserName = ""
          strPassword = ""
          strFName = ""
          strLName = ""
          intContactTypeID = Application("DEFAULT_CONTACT_TYPE")
          intOrgID = 0
          intDeptID = 0
          intLangID = Application("DEFAULT_LANGUAGE")
          strHomePhone = ""
          strOfficePhone = ""
          strMobilePhone = ""
          strJobTitle = ""
          intJobFunctionID = 0
          strEmail = ""
          strPagerEmail = ""
          strOfficeLocation = ""
          strNotes = ""
          intIOStatusID = 0
          dtIOStatusDate = ""
          strIOStatusText = ""
          intTZOffset = 0
          strResume  = ""
          intPhotoFileID = 0
          dtLastAccess = ""
          dteLastUpdate = ""
          intLastUpdateByID = 0
      						
        Case Else
          ' Do nothing
      					
      End Select

    Case 2  ' Edit Contact

      Set objContact = New clsContact

      objContact.ID = intContactID
        			
      If Not objContact.Load Then
        				
        ' Couldn't load user for some reason
        				
      Else
        				
        strUserName = objContact.UserName
        strPassword = objContact.Password
        strFName = objContact.FName
        strLName = objContact.LName
        intContactTypeID = objContact.ContactTypeID
        intOrgID = objContact.OrgID
        intDeptID = objContact.DeptID
        intLangID = objContact.LangID
        strHomePhone = objContact.HomePhone
        strOfficePhone = objContact.OfficePhone
        strMobilePhone = objContact.MobilePhone
        strJobTitle = objContact.JobTitle
        					
        If IsNull(objContact.JobFunctionID) Then
          intJobFunctionID = 0
        Else
          intJobFunctionID = objContact.JobFunctionID
        End If
        					
        strEmail = objContact.Email
        strPagerEmail = objContact.PagerEmail
        strOfficeLocation = objContact.OfficeLocation
        strNotes = objContact.Notes
        					
        intIOStatusID = objContact.IOStatusID
        dtIOStatusDate = objContact.IOStatusDate
        strIOStatusText = objContact.IOStatusText
        intTZOffset = objContact.TZOffset

        strResume  = objContact.sResume
        intPhotoFileID = objContact.PhotoFileID
        dtLastAccess = objContact.LastUpdate
        dteLastUpdate = objContact.LastUpdate
        intLastUpdateByID = objContact.LastUpdateByID

      End If

      Set objContact = Nothing

    Case Else
      ' Do nothing
    				
  End Select

%>

  <TR>
    <TD>
  	  <TABLE class=Normal cellSpacing=1 cellPadding=1 width="100%" border=0 bgColor=white>
	    <FORM action="Contact.asp" method="post" id=frmContact name=frmContact>
	    <INPUT id=tbxSave name=tbxSave type=hidden value="1">
	    <TR class="lhd_Heading1">
	    	<TD colspan=5 align=center><%=strHeader%></TD>
	    </TR>
	    <TR>
	    	<TD colspan=5></TD>
	    </TR>
      <TR>
        <TD width="22%"><B><%=Lang("User_Name")%>:</B><INPUT id=tbxContactID name=tbxContactID type=hidden value="<%=intContactID%>"></TD>
        <TD width="25%">
	      	<%
	      	' Mode = 1  New Contact
	      	'      = 2  Edit Contact
	      	'
	      	If strMode = "1" Then
	      	  If Application("AUTH_TYPE") = 1 Then
	      	  %>
	      	  	<%=strUserName%>
	      	  	<INPUT type="hidden" id=tbxUserName name=tbxUserName value="<%=strUserName%>" style="WIDTH: 100%">
	      	  <%
	      	  Else
	      	  %>
	      	  	<INPUT id=tbxUserName name=tbxUserName value="<%=strUserName%>" style="WIDTH: 100%">
	      	  <%
	      	  End If
	      	Else
	      	%>
	      		<%=strUserName%>
      	  	<INPUT type="hidden" id=tbxUserName name=tbxUserName value="<%=strUserName%>" style="WIDTH: 100%">
	      	<%
	      	End If
	      	%>
	      </TD>
        <TD width="5%"></TD>
        <TD width="18%"><%=Lang("Last_Access_Time")%>:</TD>
        <TD width="25%"><%=DisplayDateTime(dtLastAccess)%></TD>
      </TR>
      <TR>
        <TD><%=Lang("Password")%>:</TD>
        <TD><INPUT id=pwdPassword name=pwdPassword type=password style="WIDTH: 100%" value="<%=strPassword%>"></TD>
        <TD></TD>
        <TD></TD>
        <TD></TD>
      </TR>
      <TR height=27>
        <TD><B><%=Lang("First_Name")%>:</TD>
        <TD><INPUT id=tbxFName name=tbxFName style="WIDTH: 100%" value="<%=strFName%>"
             ></TD>
        <TD style="WIDTH: 50px" width=50></TD>
        <TD><B><%=Lang("Last_Name")%>:</B></TD>
        <TD><INPUT id=tbxLName name=tbxLName style="WIDTH: 100%"  value="<%=strLName%>"
             ></TD>
      </TR>
      <TR height=27>
        <TD><B><%=Lang("Email")%>:</B></TD>
        <TD><INPUT id=tbxEmail name=tbxEmail 
                style="WIDTH: 100%" value="<%=strEmail%>"></TD>
        <TD></TD>
        <TD></TD>
        <TD></TD>
      </TR>
      <TR>
        <TD align=left><%=Lang("Pager_Email")%>:</TD>
        <TD><INPUT id=tbxPagerEmail name=tbxPagerEmail style="WIDTH: 100%" value="<%=strPagerEmail%>"></TD>
        <TD></TD>
        <TD></TD>
        <TD></TD>
      </TR>
      <TR>
        <TD><%=Lang("Contact_Type")%>:</TD>
        <TD>
          <SELECT id=cbxContactType name=cbxContactType style="WIDTH: 100%">
            <OPTION SELECTED VALUE="0">(None)</OPTION>
	    	    <%
	    	    Response.Write BuildList("CONTACT_TYPE_LIST", intContactTypeID)
	    	    %>
          </SELECT>
        </TD>
        <TD></TD>
        <TD></TD>
        <TD></TD>
      </TR>
      <TR>
        <TD><%=Lang("Organisation")%>:</TD>
        <TD>
          <SELECT id=cbxOrganisation name=cbxOrganisation style="WIDTH: 100%">
            <OPTION SELECTED VALUE="0">(None)</OPTION>
            <%
            Set objCollection = New clsCollection
                          	                   
            objCollection.CollectionType = objCollection.clOrganisation
            objCollection.Query = "SELECT OrgPK, OrgName FROM tblOrganisations WHERE IsActive=" & lhd_True & " ORDER BY OrgPK ASC"
                          		                						
            If Not objCollection.Load Then
                          		
            	Response.Write objCollection.LastError
                          		                	
            Else
                          		
              Do While Not objCollection.EOF
                          		                    
             	  If objCollection.Item.ID = intOrgID Then
             		%>
             			<OPTION SELECTED VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.OrgName%></OPTION>
             		<%
             		Else
             		%>
             			<OPTION VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.OrgName%></OPTION>
             		<%
             		End If
                          		                		
             		objCollection.MoveNext
                          		                		
              Loop
                          		                    
            End If
                          		                					
            Set objCollection = Nothing
            %>
          </SELECT>
        </TD>
        <TD></TD>
        <TD></TD>
        <TD></TD>
      </TR>
      <TR>
        <TD class="lhd_Required"><%=Lang("Department")%>:</TD>
        <TD>
        <SELECT id=cbxDepartment name=cbxDepartment style="WIDTH: 100%">
	       <OPTION SELECTED VALUE="0">(None)</OPTION>
	       <%
	    	Set objCollection = New clsCollection
	       
	    	objCollection.CollectionType = objCollection.clDepartment
	    	objCollection.Query = "SELECT DeptPK, DeptName FROM tblDepartments WHERE IsActive=" & lhd_True & " ORDER BY DeptName ASC"
	    							
	    	If Not objCollection.Load Then
	    	
	    		Response.Write objCollection.LastError
	    		
	    	Else
	    	
	    	    Do While Not objCollection.EOF
	    	    
	    			If objCollection.Item.ID = intDeptID Then
	    			%>
	    				<OPTION SELECTED VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.DeptName%></OPTION>
	    			<%
	    			Else
	    			%>
	    				<OPTION VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.DeptName%></OPTION>
	    			<%
	    			End If
	    			
	    			objCollection.MoveNext
	    			
	    	    Loop
	    	    
	    	End If
	    						
	    	Set objCollection = Nothing
	    	%>
             </SELECT>
             </TD>
        <TD></TD>
        <TD><%=Lang("Location")%>:</TD>
        <TD><INPUT id=tbxOfficeLocation name=tbxOfficeLocation 
                style="WIDTH: 100%" value="<%=strOfficeLocation%>"></TD>
      </TR>
      <TR>
        <TD><%=Lang("Job_Title")%>:</TD>
        <TD><INPUT id=tbxJobTitle name=tbxJobTitle style="WIDTH: 100%" value="<%=strJobTitle%>"></TD>
        <TD></TD>
        <TD><%=Lang("Job_Function")%>:</TD>
        <TD>
          <SELECT id=cbxJobFunction name=cbxJobFunction style="WIDTH: 100%">
            <OPTION SELECTED VALUE="0">(None)</OPTION>
            <%
            Response.Write BuildList("JOB_FUNCTION_LIST", CInt(intJobFunctionID))
            %>
          </SELECT>
        </TD>
      </TR>
      <TR>
        <TD class="lhd_Required"><%=Lang("Phone_Work")%>:</TD>
        <TD><INPUT id=tbxOfficePhone name=tbxOfficePhone 
                style="WIDTH: 100%" value="<%=strOfficePhone%>"></TD>
        <TD></TD>
        <TD></TD>
        <TD></TD>
      </TR>
      <TR>
        <TD><%=Lang("Phone_Home")%>:</TD>
        <TD><INPUT id=tbxHomePhone name=tbxHomePhone 
                style="WIDTH: 100%" value="<%=strHomePhone%>"></TD>
        <TD></TD>
        <TD> </TD>
        <TD></TD>
      </TR>
      <TR>
        <TD><%=Lang("Phone_Mobile")%>:</TD>
        <TD><INPUT id=tbxMobilePhone name=tbxMobilePhone style="WIDTH: 100%" value="<%=strMobilePhone%>"></TD>
        <TD></TD>
        <TD></TD>
        <TD></TD>
      </TR>
      <TR>
        <TD><%=Lang("Language")%>:</TD>
        <TD>
          <SELECT id=cbxLanguage name=cbxLanguage style="WIDTH: 100%">
            <%
            Set objCollection = New clsCollection
                  	       
            objCollection.CollectionType = objCollection.clLanguage
            objCollection.Query = "SELECT LangPK, LangName FROM tblLanguages WHERE IsActive=" & lhd_True & " ORDER BY LangPK ASC"
                  	    							
            If Not objCollection.Load Then
                  	    	
            	Response.Write objCollection.LastError
                  	    		
            Else
                  	    	
              Do While Not objCollection.EOF
                  	    	    
            		If objCollection.Item.ID = intLangID Then
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
        <TD></TD>
        <TD></TD>
        <TD></TD>
      </TR>
      <TR>
        <TD vAlign=top><%=Lang("Resume")%>:</TD>
        <TD colspan=4><TEXTAREA id=txtResume name=txtResume style="HEIGHT: 70px; WIDTH: 100%"><%=strResume%></TEXTAREA></TD>
      </TR>
      <TR>
        <TD></TD>
        <TD></TD>
        <TD></TD>
        <TD></TD>
        <TD></TD>
      </TR>
      <TR>
        <TD><%=Lang("In/Out_Status")%>:</TD>
        <TD>
          <SELECT id=cbxIOStatus name=cbxIOStatus style="WIDTH: 100%">
       	    <OPTION value="0" selected>(None)</OPTION>
          </SELECT>
        </TD>
        <TD></TD>
        <TD><%=Lang("In/Out_Status_Date")%>:</TD>
        <TD><INPUT id=tbxIOStatusDate name=tbxIOStatusDate style="WIDTH: 100%" value="<%=dtIOStatusDate%>"></TD>
      </TR>
      <TR>
        <TD><%=Lang("In/Out_Status_Text")%>:</TD>
        <TD colspan=4><INPUT id=tbxIOStatusText name=tbxIOStatusText style="WIDTH: 100%" value="<%=strIOStatusText%>"></TD>
      </TR>
      <TR>
        <TD></TD>
        <TD></TD>
        <TD></TD>
        <TD></TD>
        <TD></TD>
      </TR>
      <TR>
        <TD><%=Lang("Timezone_Offset")%>:</TD>
        <TD>
          <SELECT id=cbxTZOffset name=cbxTZOffset style="WIDTH: 100%">
       	    <OPTION value="0" selected>(None)</OPTION>
      	  </SELECT>
        </TD>
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
      <TR>
        <TD><%=Lang("Photo_File")%>:</TD>
        <TD>
           <SELECT id=cbxPhotoFile name=cbxPhotoFile style="WIDTH: 100%">
       	    <OPTION value="0" selected>(None)</OPTION>
           </SELECT>
        </TD>
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
      <TR>
        <TD vAlign=top><%=Lang("Notes")%>:</TD>
        <TD colspan=4><TEXTAREA id=txtNotes name=txtNotes style="HEIGHT: 70px; WIDTH: 100%"><%=strNotes%></TEXTAREA></TD>
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
        <TD align=right><INPUT style="WIDTH: 80px; BACKGROUND-COLOR: white" type="submit" value="<%=Lang("Save")%>" name=btnSave id=btnSave></TD>
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