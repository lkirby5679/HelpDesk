<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: admLanguageString.asp
'  Date:     $Date: 2004/03/17 00:08:28 $
'  Version:  $Revision: 1.1 $
'  Purpose:  Administration page for creating/modifing LanguageStrings
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
<!-- #Include File = "Classes/clsFile.asp" -->
<!-- #Include File = "Classes/clsGroup.asp" -->
<!-- #Include File = "Classes/clsLanguage.asp" -->
<!-- #Include File = "Classes/clsLanguageLabel.asp" -->
<!-- #Include File = "Classes/clsLanguageText.asp" -->
<!-- #Include File = "Classes/clsListItem.asp" -->
<!-- #Include File = "Classes/clsMail.asp" -->
<!-- #Include File = "Classes/clsNote.asp" -->
<!-- #Include File = "Classes/clsOrganisation.asp" -->
<!-- #Include File = "Classes/clsParameter.asp" -->
<!-- #Include File = "Classes/clsRole.asp" -->

<HTML>
  <HEAD>
    <META content="MSHTML 6.00.2600.0" name="GENERATOR">
  </HEAD>
  <LINK rel="stylesheet" type="text/css" href="Default.css">
    <%
	Dim cnnDB
	Dim binUserPermMask, binRequiredPerm
	Dim blnSave
  Dim objLangText, objLangLabel, objLanguage
  Dim strLang, strLangText, strLangLabel, strMode, strHeading
  Dim intUserID, intLastUpdateByID, intLangID, intLangLabelID, intLangTextID
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
      <P align="center">
        <TABLE class="Normal" align="center" cellSpacing="1" cellPadding="1" width="680" border="0">
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

		intLangID = Request.Form("tbxLangID")
		intLangLabelID = Request.Form("tbxLangLabelID")
		intLangTextID = Request.Form("tbxLangTextID")
		strLangText = Request.Form("tbxLangText")
		dteLastUpdate = SQLDateFormat(Day(Now()), Month(Now()), Year(Now())) & " " & FormatDateTime(Now(), vbShortTime)
		intLastUpdateByID = intUserID

    If intLangTextID > 0 Then

      ' We are just updating a current record
      
    Else
    
      ' Need to also create a new record in tblLanguageLabels
  
  		strLangLabel = Request.Form("tbxLangLabel")
  		
  		Set objLangLabel = New clsLanguageLabel
  		
  		objLangLabel.ID = intLangLabelID
  		
  		If Not objLangLabel.Load Then
  		
  		  ' Record not found, so is ok to create a new one.
  		  
  		  objLangLabel.LangLabel = strLangLabel
		    objLangLabel.LastUpdate = dteLastUpdate
		    objLangLabel.LastUpdateByID = intLastUpdateByID

  		  objLangLabel.Update
  		  
  		  intLangLabelID = objLangLabel.ID
  		  
  		Else
  		
  		  ' Do nothing as record already exists
  		  
  		End If
  		
  		Set objLangLabel = Nothing
    
    End If

		' Check for required fields
		
		If Len(strLangText) = 0 Then
		
			Call DisplayError(1, "All required fields need to be entered, please go back and populate these fields")
			
		Else
		
			' Do nothing
			
		End If

    

		' Now save/update the LanguageString details

		Set objLangText = New clsLanguageText

		objLangText.ID = intLangTextID

		If Not objLangText.Load Then
		
			' LanguageString does not exist, so we are set to create a new one

		Else
		
			' LanguageString exists, so we update
		
		End If

		' Check the the fields and leave Null if nothing is set.

		objLangText.ID = intLangTextID
		objLangText.LangID = intLangID
		objLangText.LangLabelID = intLangLabelID
		objLangText.LangText = strLangText
		objLangText.LastUpdate = dteLastUpdate
		objLangText.LastUpdateByID = intLastUpdateByID

						
		If Not objLangText.Update Then
						
			' Failed to create/save LanguageString
			Response.Write objLangText.LastError
							
		Else
						
			intLangTextID = objLangText.ID
			strHeading = Lang("Language_String_Saved")
						
		End If
						
		Set objLangText = Nothing
%>
            <TR>
              <TD>
                <TABLE class="Normal" cellSpacing="0" cellPadding="1" width="100%" border="0" bgColor="white">
                  <TR class="lhd_Heading1">
                    <TD></TD>
                    <TD colspan="3" align="middle"><%=strHeading%></TD>
                    <TD></TD>
                  </TR>
                  <TR>
                    <TD align="right" colspan="5"><a style="FONT-SIZE: 8pt" href="admLanguageStringList.asp?Language=<%=intLangID%>&Page=1"><%=Lang("Manage_Language_Strings")%></a></TD>
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
                    <TD colspan="3" align="left">LanguageString information has been successfully saved.</TD>
                    <TD></TD>
                  </TR>
                  <TR>
                    <TD></TD>
                    <TD colspan="3"></TD>
                    <TD></TD>
                  </TR>
                </TABLE>
              </TD>
            </TR>
            <%
	Else
	
		' Mode: 1 - To create a new LanguageString
		'		2 - Edit a LanguageString

		If  Not IsEmpty(Request.QueryString("Language")) Then
	
			' Create a new record
	
			strMode = 1
			intLangTextID = 0
			strHeading = Lang("New_Language_String")
			
		Else
	
			' Edit a record determine by the LanguageStringID passed via the QueryString
	
			strMode = 2
			intLangTextID = Request.QueryString("ID")
			strHeading = Lang("Modify_Language_String")
			
		End If
			

		Select Case strMode
		
			Case 1	' Create new LanguageString
			
				intLangID = Request.QueryString("Language")
				strLang = ""
				intLangLabelID = 0
				strLangLabel = ""
				strLangText = ""
				dteLastUpdate = ""
				intLastUpdateByID = 0


			Case 2  ' Edit LanguageString

				' Get the LanguageString ID we want to edit and load the record

				Set objLangText = New clsLanguageText

				objLangText.ID = intLangTextID
			
				If Not objLangText.Load Then
				
					' Couldn't load user for some reason
					Response.Write objLangText.LastError
				
				Else
				
				  intLangID = objLangText.LangID
				  strLang = objLangText.LangID
				  intLangLabelID = objLangText.LangLabelID
				  strLangLabel = objLangText.LangLabel.LangLabel
				  strLangText = objLangText.LangText
					dteLastUpdate = objLangText.LastUpdate
					intLastUpdateByID = objLangText.LastUpdateByID

				End If

				Set objLangText = Nothing
				

			Case Else
				' Do nothing
				
		End Select

    
    ' Get the language name

    Set objLanguage = New clsLanguage
    
    objLanguage.ID = intLangID
    
    If Not objLanguage.Load Then
      ' Record not loaded
    Else
      strLang = objLanguage.LangName
    End If
    
    Set objLanguage = Nothing


%>
            <TR>
              <TD>
                <TABLE class="Normal" cellSpacing="0" cellPadding="1" width="100%" border="0" bgColor="white">
                  <FORM action="admLanguageString.asp" method="post" id="frmLanguageString" name="frmLanguageString">
                    <INPUT id="tbxLangLabelID" name=tbxLangLabelID type=hidden value="<%=intLangLabelID%>">
                    <INPUT id="tbxLangID" name=tbxLangID type=hidden value="<%=intLangID%>">
                    <INPUT id=tbxLangTextID name=tbxLangTextID type=hidden value="<%=intLangTextID%>">
                    <INPUT id="tbxSave" name="tbxSave" type="hidden" value="1">
                    <TR class="lhd_Heading1">
                      <TD colspan="5" align="center"><%=strHeading%></TD>
                    </TR>
                    <TR>
                      <TD align="right" colspan="5"><a style="FONT-SIZE: 8pt" href="admLanguageStringList.asp?Language=<%=intLangID%>&Page=1"><%=Lang("Manage_Language_Strings")%></a></TD>
                    </TR>
                    <TR>
                      <TD></TD>
                      <TD></TD>
                      <TD></TD>
                      <TD></TD>
                      <TD></TD>
                    </TR>
                    <TR>
                      <TD width="22%"><b><%=Lang("Language_Label")%>:</b></TD>
                      <TD width="25%">
                        <%
                        If strMode = 1 Then
                        %>
                          <INPUT id="tbxLangLabel" name=tbxLangLabel style="WIDTH: 100%" value="<%=strLangLabel%>">
                        <%
                        Else
                          Response.Write strLangLabel
                        End If
                        %>
                      </TD>
                      <TD width="5%"></TD>
                      <TD width="18%"><%=Lang("Language")%>:</TD>
                      <TD width="25%"><%=strLang%></TD>
                    </TR>
                    <TR>
                      <TD><B><%=Lang("Language_Text")%>:</B></TD>
                      <TD><INPUT id=tbxLangText name=tbxLangText style="WIDTH: 100%" value="<%=strLangText%>"></TD>
                      <TD></TD>
                      <TD></TD>
                      <TD></TD>
                    </tr>
                    <TR>
                      <TD></TD>
                      <TD></TD>
                      <TD></TD>
                      <TD></TD>
                      <TD align="right"><INPUT style="WIDTH: 80px; BACKGROUND-COLOR: white" id=btnSave name=btnSave type=submit value="<%=Lang("Save")%>"></TD>
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
    </BODY></HTML>
<%
  
cnnDB.Close
Set cnnDB = Nothing
  
%>
