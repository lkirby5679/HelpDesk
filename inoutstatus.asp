<% Option Explicit %>

<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: Contact.asp
'  Date:     $Date: 2004/03/29 08:23:52 $
'  Version:  $Revision: 1.1 $
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

Dim strFName, strLName, strIOStatusText

Dim intUserID, intLastUpdateByID
Dim intContactID, intIOStatusID


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
  strFName = Request.Form("tbxFName")
  strLName = Request.Form("tbxLName")
  intIOStatusID = CInt(Request.Form("cbxIOStatus"))
'  dtIOStatusDate = SQLDate( Request.Form("tbxIOStatusDate") )
  dtIOStatusDate = SQLDateFormat(Day(Now()), Month(Now()), Year(Now())) & " " & FormatDateTime(Now(), vbShortTime)
  '** If status = in then clear text
  strIOStatusText = Request.Form("tbxIOStatusText")
  dteLastUpdate = SQLDateFormat(Day(Now()), Month(Now()), Year(Now())) & " " & FormatDateTime(Now(), vbShortTime)
  intLastUpdateByID = intUserID
  ' Now save the status details

  Set objContact = New clsContact

  objContact.ID = intContactID

  If Not objContact.Load Then
    ' Contact does not exist
  Else
    ' Contact exists
  End If

  If intIOStatusID > 0 Then
    objContact.IOStatusID = intIOStatusID
  End If

  If IsDate(dtIOStatusDate) Then
    objContact.IOStatusDate = dtIOStatusDate
  End If

  objContact.IOStatusText = strIOStatusText

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
        <TR class="lhd_Heading1"><TD colspan=5 align=center><%=Lang("Saved")%></TD></TR>
        <TR><TD colspan=5 align=center></TD></TR>
        <TR><TD colspan=5 align=center><A href="inoutStatus.asp?id=<%=intContactID%>"><%=strFName & " " & strLName%></A>&nbsp;<%=Lang("Status Saved")%></TD></TR>
        <TR><TD colspan=5 align=center></TD></TR>
        <tr><td colspan=5 align=right><a href="inoutList.asp"><%=lang("IO_Board")%></a></td></tr>
      </TABLE>
   </TD>
</TR>
<%
Else
  intContactID = Request.QueryString("id")
  Set objContact = New clsContact
  objContact.ID = intContactID
  If Not objContact.Load Then
    ' Couldn't load user for some reason
  Else
    strFName = objContact.FName
    strLName = objContact.LName
    intIOStatusID = objContact.IOStatusID
    dtIOStatusDate = objContact.IOStatusDate
    strIOStatusText = objContact.IOStatusText
    dteLastUpdate = objContact.LastUpdate
    intLastUpdateByID = objContact.LastUpdateByID
  End If
%>
  <TR>
    <TD>
      <TABLE class=Normal cellSpacing=1 cellPadding=1 width="100%" border=0 bgColor=white>
      <FORM action="inoutStatus.asp" method="post" id=frmStatus name=frmStatus>
      <INPUT id=tbxSave name=tbxSave type=hidden value="1">
      <INPUT id=tbxContactID name=tbxContactID type=hidden value="<%=intContactID%>">
      <INPUT id=tbxFName name=tbxFName type=hidden value="<%=strFName%>">
      <INPUT id=tbxLName name=tbxLName type=hidden value="<%=strLName%>">
      <INPUT id=tbxIOStatusDate name=tbxIOStatusDate type=hidden value="<%=dtIOStatusDate%>">
      <TR class="lhd_Heading1"><TD colspan=2 align=center><%=lang("Status_information")%></TD></TR>
      <TR><TD colspan=2></TD></TR>
      <TR>
      	<TD><%=lang("Name")%></TD>
      	<TD><%=strFName & " " & strLName%></TD>
     </TR>
      <TR>
        <TD><%=Lang("In/Out_Status")%>:</TD>
        <TD>
          <SELECT id=cbxIOStatus name=cbxIOStatus>
       	  <%
   	       Response.Write BuildList("INOUT_LIST_STATUS", intIOStatusID)
   	  %>
          </SELECT>
        </TD>
      </TR>
      <TR>
        <TD width=20%><%=Lang("In/Out_Status_Text")%>:</TD>
        <TD><INPUT id=tbxIOStatusText name=tbxIOStatusText style="WIDTH: 100%" value="<%=strIOStatusText%>"></TD>
      </TR>
      <TR>
        <TD colspan=2 align=right><INPUT style="WIDTH: 80px; BACKGROUND-COLOR: white" type="submit" value="<%=Lang("Save")%>" name=btnSave id=btnSave></TD>
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