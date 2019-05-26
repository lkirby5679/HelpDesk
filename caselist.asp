<%@ LANGUAGE="VBScript" %>
<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: caseList.asp
'  Date:     $Date: 2004/03/11 06:08:42 $
'  Version:  $Revision: 1.4 $
'  Purpose:  Lists cases
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
Dim objCollection, objRep, objGroups
Dim binUserPermMask
Dim rsCaseTypes
Dim strHTML, strMode, strHeading, strNoOfCases
Dim I, intUserID, intRepID, intPage, intPages
Dim strORDERBY, strSQL
Dim strColumn, strColumnOrder



' Create the connection to the database
Set cnnDB = CreateConnection

' Determine the logged in User's ID
intUserID = GetUserID
binUserPermMask = GetUserPermMask


' Get variable from the QueryString
intPage = CInt(Request.Querystring("Page"))
strMode = Request.QueryString("Mode")
strColumn = Request.QueryString("Column")
strColumnOrder = Request.QueryString("Order")

	
If intPage = 0 Then
	intPage = 1
Else
	' Do nothing
End If
	

' Listed cases will depend on who is logged in and what access permission they have. At
' the moment I am just listing all the cases.
'
' Mode  1.) List all Cases , that are not "Closed", and are raised by the logged
'			on user, including those that the user is included on the Case Cc
'			list
'
'		2.) List all Unassigned Cases for Rep logged in.  This includes all cases that
'			do not have a Rep assigned and that the Rep logged in belongs to the Group
'			assigned to the Case, and remembering to disregard "Closed" Cases.
'			NOTE: Exception to this rule is those case that DONT have Rep and/or Group
'			assigned
'
'		3.) List all the Cases assigned to the logged in user that are not "Closed"
'
'		4.) List assigned Case's for the given Rep
	
	
Select Case strMode
		
	Case "1"
		
		strHeading = "My Active Cases"
		
		strSQL = "SELECT * FROM tblCases " &_
				     "WHERE (IsActive=" & lhd_True & ") " &_
						   "AND ((ContactFK=" & intUserID & ") OR (Cc LIKE '%" & Session("lhd_UserName") & "%')) " &_
						   "AND (StatusFK<>" & Application("STATUS_CLOSED") & ") " &_
						   "AND (StatusFK<>" & Application("STATUS_CANCELLED") & ")"

	Case "2"
		
		strHeading = "Unassigned Cases"

		' First determine all group the Rep belongs to and then match these to all unassigned cases
		strSQL = "SELECT tblGroups.* FROM tblGroups " & _
				     "INNER JOIN tblGroupMembers ON tblGroups.GroupPK=tblGroupMembers.GroupFK " & _
				     "WHERE (tblGroupMembers.ContactFK=" & intUserID & ") " &_
					     "AND (tblGroups.IsActive=" & lhd_True & ") " &_
				     "ORDER BY tblGroups.GroupPK ASC"
					 
		Set objGroups = New clsCollection

		objGroups.CollectionType = objGroups.clGroup
		objGroups.Query = strSQL

		If Not objGroups.Load Then
				
			' Raise Error
					
		Else
			
			If objGroups.BOF And objGroups.EOF Then
				
				' No records
				strSQL = ""
					
			Else
			
				strSQL = "SELECT tblCases.* FROM tblCases "
				strSQL = strSQL & "WHERE (tblCases.IsActive=" & lhd_True & ") "
 				strSQL = strSQL & "AND (tblCases.RepFK Is Null) "
				strSQL = strSQL & "AND (tblCases.StatusFK<>" & Application("STATUS_CLOSED") & ") "
				strSQL = strSQL & "AND (tblCases.StatusFK<>" & Application("STATUS_CANCELLED") & ") "
 				strSQL = strSQL & "AND ((tblCases.GroupFK=" & objGroups.Item.ID & ") "

				While Not objGroups.EOF

 					strSQL = strSQL & "OR (tblCases.GroupFK=" & objGroups.Item.ID & ") "
					
					objGroups.MoveNext
				
				WEnd

				strSQL = strSQL & ")"
			
			End If
			
		End If
		
		Set objGroups = Nothing
				 
		
	Case "3"
		
		strHeading = Session("lhd_Username") & "'s Assigned Cases"

		strSQL = "SELECT * FROM tblCases " &_
    				 "WHERE (IsActive=" & lhd_True & ") " &_ 
		     	     "AND (RepFK = " & intUserID & ") " &_
						   "AND (StatusFK<>" & Application("STATUS_CLOSED") & ") " & _
						   "AND (StatusFK<>" & Application("STATUS_CANCELLED") & ")"

	Case "4"
		
		If Request.QueryString("rep") > 0 Then
			intRepID = CInt(Request.QueryString("rep"))
		Else
			intRepID = CInt(Request.Form("cbxRep"))
		End If
		
		Set objRep = New clsContact
			
		objRep.ID = intRepID
		
		If Not objRep.Load Then

			' Raise Error, contact details didn't load

		Else

			strHeading = objRep.Username & "'s Assigned Cases"
			
			strSQL = "SELECT * FROM tblCases " &_
					     "WHERE (IsActive=" & lhd_True & ") " &_ 
							    "AND (RepFK=" & objRep.ID & ") " &_
							    "AND (StatusFK<>" & Application("STATUS_CLOSED") & ") " & _
							    "AND (StatusFK<>" & Application("STATUS_CANCELLED") & ")"
						 
		End If
			
		Set objRep = Nothing
		
	Case Else
	
		' Catch any alternative
		
End Select

' Build the ORDER BY string
			
Select Case strColumn
		
  Case "1"
    strORDERBY = "ORDER BY CasePK"
		  
  Case "2"
    strORDERBY = "ORDER BY Title"
		  
  Case "3"
    strORDERBY = "ORDER BY RepFK"
		  
  Case "4"
    strORDERBY = "ORDER BY StatusFK"
		  
  Case "5"
    strORDERBY = "ORDER BY RaisedDate"
		  
  Case Else
    strORDERBY = "ORDER BY CasePK"
		
End Select
		
If strColumnOrder = "1" Then
  strORDERBY = strORDERBY & " DESC"
Else
  strORDERBY = strORDERBY & " ASC"
End If


' Generate the query and build the results table

If Len(strSQL) > 0 Then

  strSQL = strSQL & " " & strORDERBY

  Set objCollection = New clsCollection
  	
  objCollection.CollectionType = objCollection.clCase
  objCollection.Query = strSQL

  If Not objCollection.Load Then
  	
  	' Raise Error
  		
  Else
  	
  	If objCollection.BOF And objCollection.EOF Then
  		
  		' No records returned
  			
  	Else
  		
      strNoOfCases = CStr(objCollection.RecordCount)
  	
  		If objCollection.RecordCount Mod Application("ITEMS_PER_PAGE") = 0 Then
  			intPages = Int(objCollection.RecordCount / Application("ITEMS_PER_PAGE"))
  		Else
  			intPages = Int(objCollection.RecordCount / Application("ITEMS_PER_PAGE")) + 1
  		End If
  		
  		strHTML = ""
  		
  		' Move the the record at the start of the next intPage
  		objCollection.Move(Application("ITEMS_PER_PAGE") * (intPage - 1))

  		I = 0
  	
  		Do While Not objCollection.EOF And Application("ITEMS_PER_PAGE") > I
  		
  		  ' Alternate row background colours to make it easier to read
  		  If I Mod 2 > 0 Then
  		    ' Odd row number
    			strHTML = strHTML & "<TR bgcolor=""white"" class=""lhd_TableRow_Odd"">" & Chr(13)
  		  Else
  		    ' Even row number
    			strHTML = strHTML & "<TR bgcolor=""WhiteSmoke"" class=""lhd_TableRow_Even"">" & Chr(13)
  		  End If
  		  
  			strHTML = strHTML & "	<TD align=""Center"">" & objCollection.Item.ID & "</TD>" & Chr(13)
  			strHTML = strHTML & "	<TD>&nbsp;<A href=""caseModify.asp?ID=" & objCollection.Item.ID & """>" & objCollection.Item.Title & "</A></TD>" & Chr(13)

  			If IsEmpty(objCollection.Item.Rep.UserName) Then
  				strHTML = strHTML & "	<TD align=""Center"">-</TD>" & Chr(13)
  			Else
  				strHTML = strHTML & "	<TD align=""Center"">" & objCollection.Item.Rep.UserName & "</TD>" & Chr(13)
  			End If

  			strHTML = strHTML & "	<TD align=""Center"">" & objCollection.Item.Status.ItemName & "</TD>" & Chr(13)
  			strHTML = strHTML & "	<TD>&nbsp;" & DisplayDateTime(objCollection.Item.RaisedDate)& "</TD>" & Chr(13)
  			strHTML = strHTML & "</TR>" & Chr(13)
  				
  			I = I + 1
  				
  			objCollection.MoveNext
  		Loop
  			
  	End If
  		
  End If

  Set objCollection = Nothing

Else

  ' Do nothing

End If
	
%>

<HTML>

<HEAD>
	
	<META content="MstrHTML 6.00.2600.0" name=GENERATOR>
</HEAD>

<LINK rel="stylesheet" type="text/css" href="default.css">

<BODY bgColor=WhiteSmoke>
<P align=center>
<TABLE align=center style="WIDTH: 680px" cellSpacing=1 cellPadding=1 width="680" border=0>
	<TR>
		<TD>
			<%
			Response.Write DisplayHeader
			%>
		</TD>
	</TR>
	<TR>
		<TD>
			<TABLE class="lhd_Table_Normal" cellspacing="1" cellpadding="0">
				<TR class="lhd_Heading1">
					<TD colspan="5" align="Centre">&nbsp;<%=strHeading%>&nbsp;(<%=strNoOfCases%>)</TD>
				</TR>
				<TR>
					<TH width="8%" align="Center"><A href="caseList.asp?Mode=<%=strMode%>&Page=<%=intPage%>&Column=1&Order=<%If strColumnOrder=1 Then Response.Write "0" Else Response.Write "1" End If%>"><%=Lang("Case")%> #</A></TH>
					<TH width="43%" align="Left">&nbsp;<A href="caseList.asp?Mode=<%=strMode%>&Page=<%=intPage%>&Column=2&Order=<%If strColumnOrder=1 Then Response.Write "0" Else Response.Write "1" End If%>"><%=Lang("Title")%></A></TD>
					<TH width="12%" align="Center"><A href="caseList.asp?Mode=<%=strMode%>&Page=<%=intPage%>&Column=3&Order=<%If strColumnOrder=1 Then Response.Write "0" Else Response.Write "1" End If%>"><%=Lang("Assigned")%></A></TH>
					<TH width="15%" align="Center"><A href="caseList.asp?Mode=<%=strMode%>&Page=<%=intPage%>&Column=4&Order=<%If strColumnOrder=1 Then Response.Write "0" Else Response.Write "1" End If%>"><%=Lang("Status")%></A></TH>
					<TH width="17%" align="Left">&nbsp;<A href="caseList.asp?Mode=<%=strMode%>&Page=<%=intPage%>&Column=5&Order=<%If strColumnOrder=1 Then Response.Write "0" Else Response.Write "1" End If%>"><%=Lang("Start_Date")%></A></TH>
				</TR>
				<%
				Response.Write strHTML
				%>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD>
			<TABLE class=Normal cellSpacing=1 cellPadding=1 width="100%" border=0 bgColor=white>
			  <%
			  If strMode = 4 Then
			  	Response.Write DisplayPageNumbers( "caseList.asp?Mode=" & strMode & "&Rep=" & intRepID & "&Column=" & strColumn & "&Order=" & strColumnOrder & "&", intPage, intPages )
			  Else
			  	Response.Write DisplayPageNumbers( "caseList.asp?Mode=" & strMode & "&Column=" & strColumn & "&Order=" & strColumnOrder & "&", intPage, intPages )
			  End If
			  %>
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
