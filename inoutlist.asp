<%@ LANGUAGE="VBScript" %>
<%
' ----------------------------------------------------------------------------------
'  Huron Support Desk, Copyright (C) 2003
'  Huron Support Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Author(s):t_klose@hotmail.com, klunde@hotmail.com
'  Filename: inoutList.asp
'  Date:     $Date: 2004/03/29 08:23:52 $
'  Version:  $Revision: 1.6 $
'  Purpose:  This is the main menu for the In/Out Board.
' ----------------------------------------------------------------------------------
%>
<%

Option Explicit

%>
<!--METADATA TYPE="TypeLib" NAME="Microsoft ActiveX Data Objects 2.5 Library" UUID="{00000205-0000-0010-8000-00AA006D2EA4}" VERSION="2.5"-->
<!--METADATA TYPE="TypeLib" NAME="Microsoft Scripting Runtime" UUID="{420B2830-E718-11CF-893D-00A0C9054228}" VERSION="1.0"-->
<!--METADATA TYPE="TypeLib" NAME="Microsoft CDO for Windows 2000 Library" UUID="{CD000000-8B95-11D1-82DB-00C04FB1625D}" VERSION="1.0"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

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

<%
  Dim cnnDB
  Dim objCollection
  Dim I, intPages, intPage, intUserID, intPinNumber
  Dim strHTML
  Dim strORDERBY, strSQL
  Dim strColumn, strColumnOrder
  Dim binUserPermMask
  
  Dim mFirstname, mLastname, mInoutStatus, mPhone, mUid, mDept, check
  Dim strFilter, strSqlWhere
  check = ""
  
  ' Get user variables
  Set cnnDB = CreateConnection
  intUserID = GetUserID
  binUserPermMask = GetUserPermMask
  
  ' Get the settings from the QueryString
  strColumn = Request.QueryString("Column")
  strColumnOrder = Request.QueryString("Order")
  intPage = CInt(Request.Querystring("Page"))
  If intPage = 0 Then
    intPage = 1
  Else
    ' Do nothing
  End If

  if request.form("button") = lang("Search") then
    mFirstname = request.form("mFirstname")
    mLastname= request.form("mLastname")
    mUid = request.form("mUid")
    mDept = request.form("mDept")
    mPhone = request.form("mPhone")
    mInoutStatus = request.form("mInoutStatus")
  else
    mFirstname = Request.Querystring("mFirstname")
    mLastname = Request.Querystring("mLastname")
    mUid = Request.Querystring("mUid")
    mDept = Request.Querystring("mDept")
    mPhone = Request.Querystring("mPhone")
    mInoutStatus = request.Querystring("mInoutStatus")
  end if

  'Build filter string
  if len(mInoutStatus) < 1 then mInoutstatus = 0
  strFilter = "&mFirstname=" & mFirstname & _
              "&mLastname=" & mLastname & _
              "&mInoutStatus=" & mInoutStatus & _
              "&mPhone=" & mPhone & _
              "&mUid=" & mUid & _
              "&mDept=" & mDept

    if len(mFirstname) > 0 then
    	strSqlWhere = strSqlWhere & " AND (FName like '" & mFirstname & "%')"
    end if
    if len(mLastname) > 0 then
    	strSqlWhere = strSqlWhere & " AND (LName like '" & mLastname & "%')"
    end if
    if len(mUid) > 0 then
    	strSqlWhere = strSqlWhere & " AND (Username like '" & mUid & "%')"
    end if
    if len(mPhone) > 0 then
    	strSqlWhere = strSqlWhere & " AND (OfficePhone like '" & mPhone & "%')"
    end if
    if mInoutStatus = 1 then
    	check = "checked"
    	strSqlWhere = strSqlWhere & " AND (IOStatusFK >= 1)" '** Change
    end if
    if len(mDept) > 0 then
    	strSqlWhere = strSqlWhere & " AND (DeptFK like '" & mDept & "%')" '** Change
    end if

  ' Build the ORDER BY string
  Select Case strColumn
    Case "1"
      strORDERBY = "ORDER BY FName, LName"
    Case "2"
      strORDERBY = "ORDER BY LName, FName"
    Case "3"
      strORDERBY = "ORDER BY IOStatusFK, FName, LName"
    Case "4"
      strORDERBY = "ORDER BY OfficePhone, FName, LName"
    Case "5"
      strORDERBY = "ORDER BY Username, FName, LName"
    Case "6"
      strORDERBY = "ORDER BY DeptFK, FName, LName" '** Change
    Case Else
      strORDERBY = "ORDER BY FName, LName"
  End Select

  If strColumnOrder = "1" Then
    strORDERBY = strORDERBY & " DESC"
  Else
    strORDERBY = strORDERBY & " ASC"
  End If
  ' Build the list of departments

  strSQL = "SELECT ContactPK, DeptFK, FName, LName, Username, IsActive, OfficePhone,IOStatusFK, IOStatusDate, IOStatusText FROM tblContacts"
  strSQL = strSQL & " WHERE ContactPK > 0" & strSqlWhere & " " & strORDERBY
  Set objCollection = New clsCollection

  objCollection.CollectionType = objCollection.clContact
  objCollection.Query = strSQL

  If Not objCollection.Load Then
    Response.Write objCollection.LastError
  Else
    If objCollection.BOF And objCollection.EOF Then
      ' No records
    Else
      If objCollection.RecordCount Mod Application("ITEMS_PER_PAGE") = 0 Then
        intPages = Int(objCollection.RecordCount / Application("ITEMS_PER_PAGE"))
      Else
        intPages = Int(objCollection.RecordCount / Application("ITEMS_PER_PAGE")) + 1
      End If
      strHTML = ""
      ' Move the the record at the start of the next page
      objCollection.Move(Application("ITEMS_PER_PAGE") * (intPage - 1))
      I = 0
      Do While Not objCollection.EOF And Application("ITEMS_PER_PAGE") > I
        I = I + 1
        ' Make sure pinNumber is not null
        if objCollection.Item.IOStatusID then
          intPinNumber = objCollection.Item.IOStatusID - 16 '** Change
        else
          intPinNumber = 1
        end if
        ' Alternate row background colours to make it easier to read
        If I Mod 2 > 0 Then
          strHTML = strHTML & "<TR class=""inout_TableRow_Odd"">" & Chr(13)
        Else
          strHTML = strHTML & "<TR class=""inout_TableRow_Even"">" & Chr(13)
        End If
        strHTML = strHTML & "<TD class=inoutLeftColumn align=left>" & objCollection.Item.FName & "</TD>" & Chr(13)
        strHTML = strHTML & "<TD class=inoutLeftColumn align=left>" & objCollection.Item.LName & "</TD>" & Chr(13)
        strHTML = strHTML & "<TD class=inoutLeftColumn align=center><a href=""inoutContact.asp?id=" & objCollection.Item.ID & """><img src=""Images/pin" & intPinNumber & ".gif"" border="""" alt=""" & objCollection.Item.IOStatusText & " " & lang("ClickForMoreDetails") & """></a></TD>" & Chr(13)
        strHTML = strHTML & "<TD class=inoutLeftColumn align=left>" & objCollection.Item.OfficePhone & "</TD>" & Chr(13)
        strHTML = strHTML & "<TD class=inoutLeftColumn align=left>" & objCollection.Item.Username & "</TD>" & Chr(13)
        strHTML = strHTML & "<TD class=inoutLeftColumn align=left>" & objCollection.Item.Dept.DeptName & "</TD>" & Chr(13)
        If PERM_ACCESS_ADMIN = (PERM_ACCESS_ADMIN And binUserPermMask) Then
		' Admin access granted
        	strHTML = strHTML & "<TD class=inoutRightColumn align=center><A href=""admContact.asp?Mode=3&ID=" & objCollection.Item.ID & """>" & "<IMG src=""Images/Pencil.gif"" alt=""" & Lang("Edit") & """ border=""0""></A></TD>" & Chr(13)
	Else
		' User access
        	strHTML = strHTML & "<TD class=inoutRightColumn align=center>&nbsp;</TD>" & Chr(13)
	End If
       strHTML = strHTML & "</TR>" & Chr(13)
        objCollection.MoveNext
      Loop
    End If
  End If
  Set objCollection = Nothing
%>

<HEAD>
  
  <META content="MSHTML 6.00.2600.0" name=GENERATOR>
  <LINK rel="stylesheet" type="text/css" href="Default.css">
</HEAD>
<BODY>
  <P align=center>
    <TABLE align=center cellSpacing=1 cellPadding=1 width="680" border=0>
      <TR>
        <TD><% Response.Write DisplayHeader %></TD>
      </TR>
      <TR>
        <TD>
        <table class="inout">
          <form method="post" action="inoutList.asp">
          <TR class="lhd_Heading1">
            <TD colspan=7 align=center><%=Lang("IO_Board")%></TD>
          </TR>
          <tr class="normal">
          <td colspan="7" align="right">
  	    <input type="submit" name="button" value="<%=lang("Search")%>">
  	    <input type="submit" name="button" value="<%=lang("ClearForm")%>">
  	    <input type="submit" name="button" value="<%=lang("ShowAll")%>">
  	  </td></tr>
          <TR class=inoutHeading>
            <TD class=inoutLeftColumn align=left width="20%"><A href="InOutList.asp?Column=1<%=strFilter%>&Order=<%If strColumnOrder=1 Then Response.Write "0" Else Response.Write "1" End If%>"><%=Lang("First_Name")%></a></TD>
            <TD class=inoutLeftColumn align=left width="20%"><A href="InOutList.asp?Column=2<%=strFilter%>&Order=<%If strColumnOrder=1 Then Response.Write "0" Else Response.Write "1" End If%>"><%=Lang("Last_Name")%></a></TD>
            <TD class=inoutLeftColumn align=center width="10%"><A href="InOutList.asp?Column=3<%=strFilter%>&Order=<%If strColumnOrder=1 Then Response.Write "0" Else Response.Write "1" End If%>"><%=Lang("Status")%></a></TD>
            <TD class=inoutLeftColumn align=left width="10%"><A href="InOutList.asp?Column=4<%=strFilter%>&Order=<%If strColumnOrder=1 Then Response.Write "0" Else Response.Write "1" End If%>"><%=Lang("Phone")%></a></TD>
            <TD class=inoutLeftColumn align=left width="10%"><A href="InOutList.asp?Column=5<%=strFilter%>&Order=<%If strColumnOrder=1 Then Response.Write "0" Else Response.Write "1" End If%>"><%=Lang("User_Name")%></a></TD>
            <TD class=inoutLeftColumn align=left width="20%"><A href="InOutList.asp?Column=6<%=strFilter%>&Order=<%If strColumnOrder=1 Then Response.Write "0" Else Response.Write "1" End If%>"><%=Lang("Department")%></a></TD>
            <TD class=inoutRightColumn align=center width="10%"><%=Lang("Options")%></TD>
          </TR>
          <tr class="head2" align="center">
            <td><input onFocus="this.value=''" type="text"   name="mFirstname" size=16 value="<% = mFirstname %>"></td>
            <td><input onFocus="this.value=''" type="text"   name="mLastname" size=16 value="<% = mLastname %>"></td>
            <td><input type="checkbox" value=1 <% = check %> name="mInoutStatus" ></td>
            <td><input onFocus="this.value=''" type="text"   name="mPhone" size=9 value="<% = mPhone %>"></td>
            <td><input onFocus="this.value=''" type="text"   name="mUid" size=8 value="<% = mUid %>"></td>
            <td><input onFocus="this.value=''" type="text"   name="mDept" size=16 value="<% = mDept %>"></td>
            <td>&nbsp;</td>
          </tr>
          </form>
         <% Response.Write strHTML %>
        </TABLE>
      </TD>
    </TR>
    <tr>
      <td class="lhd_Body" align=right><a href="InOutPhoneList.asp"><%=Lang("Phone_List")%></a></td>
    </tr>
    <TR>
      <TD><% Response.Write DisplayFooter %></TD>
    </TR>
    </TABLE>
  </P>
</BODY>
</HTML>

<%
  cnnDB.Close
  Set cnnDB = Nothing
%>