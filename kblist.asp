<%@ LANGUAGE="VBScript" %>
<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: kbList.asp
'  Date:     $Date: 2004/03/11 06:08:42 $
'  Version:  $Revision: 1.4 $
'  Purpose:  Produces a list of knowledgebase record that match the search criteria
' ----------------------------------------------------------------------------------
%>
<%

Option Explicit

%>
<HTML>

<!--METADATA TYPE="TypeLib" NAME="Microsoft ActiveX Data Objects 2.5 Library" UUID="{00000205-0000-0010-8000-00AA006D2EA4}" VERSION="2.5"-->
<!--METADATA TYPE="TypeLib" NAME="Microsoft Scripting Runtime" UUID="{420B2830-E718-11CF-893D-00A0C9054228}" VERSION="1.0"-->
<!--METADATA TYPE="TypeLib" NAME="Microsoft CDO for Windows 2000 Library" UUID="{CD000000-8B95-11D1-82DB-00C04FB1625D}" VERSION="1.0"-->

<LINK rel="stylesheet" type="text/css" href="Default.css">

<!-- #Include File = "Include/Settings.asp" -->
<!-- #Include File = "Include/Public.asp" -->

<!-- #Include File = "Classes/clsContact.asp" -->
<!-- #Include File = "Classes/clsCollection.asp" -->
<!-- #Include File = "Classes/clsKnowledgebase.asp" -->


<%

	Dim cnnDB
	Dim objCollection
	Dim binUserPermMask
	Dim strHTML, strSQL, strWHERE, strORDERBY, strKeywords, strNoOfResults
	Dim I, intUserID, intPage, intPages
	Dim intStatusID, intPriorityID, intCaseTypeID, intCategoryID, intContactID, intRepID
	Dim dtDateFrom, dtDateTo
	Dim strColumn, strColumnOrder, strQuery



	' Create the connection to the database
	Set cnnDB = CreateConnection

	' Determine the logged in User's ID
	intUserID = GetUserID
	binUserPermMask = GetUserPermMask

  ' Check if the user has rights to view the knowledgebase
  If (PERM_KB_READ = (PERM_KB_READ And binUserPermMask)) Or (PERM_KB_MODIFY = (PERM_KB_MODIFY And binUserPermMask)) Or (PERM_KB_CREATE = (PERM_KB_CREATE And binUserPermMask)) Then

  Else
  	DisplayError 4, ""
  End If

	
	' Get the settings from the QueryString & Posted form
	intPage = CInt(Request.Querystring("Page"))
	If intPage = 0 Then
		intPage = 1
	Else
		' Do nothing
	End If


  If Len(Request.QueryString("Keywords")) > 0 Then
	  strKeywords = Request.QueryString("Keywords")
	Else
	  strKeywords = Request.Form("tbxKeywords")
	End If


	If Len(strKeywords) > 0 Then
    strSQL = "SELECT * " &_
             "FROM tblKnowledgebase " &_
             "WHERE (Issue LIKE '%" & strKeywords & "%' " &_
               "OR Cause LIKE '%" & strKeywords & "%' " &_
               "OR Resolution LIKE '%" & strKeywords & "%') " &_
               "AND IsActive=" & lhd_True & " " &_
             "ORDER BY KnowledgebasePK ASC"
             
  Else
    strSQL = "SELECT * " &_
             "FROM tblKnowledgebase " &_
             "WHERE IsActive=" & lhd_True & " " &_
             "ORDER BY KnowledgebasePK ASC"
  
  End If
	

	Set objCollection = New clsCollection
	
	objCollection.CollectionType = objCollection.clKnowledgebase
	objCollection.Query = strSQL

	If Not objCollection.Load Then
	
		Response.Write objCollection.LastError
		
	Else
	
		If objCollection.BOF And objCollection.EOF Then
		
			' No records returned
			
		Else
		  
		  strNoOfResults = objCollection.RecordCount
		  
			If objCollection.RecordCount Mod Application("ITEMS_PER_PAGE")= 0 Then
				intPages = Int(objCollection.RecordCount / Application("ITEMS_PER_PAGE"))
			Else
				intPages = Int(objCollection.RecordCount / Application("ITEMS_PER_PAGE")) + 1
			End If
		
			strHTML = ""
		
			' Move the the record at the start of the next intPage
			objCollection.Move(Application("ITEMS_PER_PAGE") * (intPage - 1))

			I = 0
	
			Do While Not objCollection.EOF And Application("ITEMS_PER_PAGE") > I

        ' Alternate row colours
		    If I Mod 2 > 0 Then
		      ' Odd row number
  		  	strHTML = strHTML & "<TR bgcolor=""white"" class=""lhd_TableRow_Odd"">" & Chr(13)
		    Else
		      ' Even row number
  		  	strHTML = strHTML & "<TR bgcolor=""WhiteSmoke"" class=""lhd_TableRow_Even"">" & Chr(13)
		    End If

				strHTML = strHTML & "	<TD align=""Center""><A href=""kbView.asp?ID=" & objCollection.Item.ID & """>" & "KB" & Right("00000000" & CStr(objCollection.Item.ID), 8) & "</A></TD>" & Chr(13)
        If Len(objCollection.Item.Issue) > 128 Then
				  strHTML = strHTML & "	<TD>&nbsp;" & Left(objCollection.Item.Issue, 124) & " ..." & "</TD>" & Chr(13)
				Else
				  strHTML = strHTML & "	<TD>&nbsp;" & objCollection.Item.Issue & "</TD>" & Chr(13)
        End If				
				strHTML = strHTML & "	<TD>&nbsp;" & DisplayDateTime(objCollection.Item.LastUpdate) & "</TD>" & Chr(13)
				strHTML = strHTML & "</TR>" & Chr(13)
				
				I = I + 1
				
				objCollection.MoveNext
			Loop
			
		End If
		
	End If
	
	Set objCollection = Nothing
	
%>

<HEAD>
  
  <META content="MstrHTML 6.00.2600.0" name=GENERATOR>
</HEAD>

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
  <TR>
    <TD>
		  <TABLE class="lhd_Table_Normal" cellSpacing="1">
		    <TR class="lhd_Heading1">
		    	<TD colspan="6" align="Center"><%=Lang("Search_Results")%>&nbsp;(<%=strNoOfResults%>)</TD>
		    </TR>
        <TR>
		      <TH width="15%" align="Center"><%=Lang("Reference_ID")%></TH>
		      <TH width="65%" align="Left">&nbsp;<%=Lang("Issue")%></TH>
		      <TH width="20%" align="Left">&nbsp;<%=Lang("Last_Updated")%></TH>
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
			strHTML = ""
			
			If intPages > 1 Then

				strHTML = strHTML & "<TR style=""FONT-WEIGHT: Bold"">"
				strHTML = strHTML & "  <TD align=""Center"" style=""FONT-SIZE: 9.5pt"">"

				If intPage > 1 Then
          strHTML = strHTML & "    <A href=""kbList.asp?" & strQuery & "&Page=" & CStr(intPage-1) & "&Column=" & strColumn & "&Order=" & strColumnOrder & """>" & Lang("Previous") & "</A>"
				Else
					strHTML = strHTML & "    <FONT color=""gray"">" & Lang("Previous") & "</FONT>"
				End If

    		strHTML = strHTML & "&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;"
    		
				If intPages > intPage Then
          strHTML = strHTML & "    <A href=""kbList.asp?" & strQuery & "&Page=" & CStr(intPage+1) & "&Column=" & strColumn & "&Order=" & strColumnOrder & """>" & Lang("Next") & "</A>"
				Else
					strHTML = strHTML & "    <FONT color=""gray"">" & Lang("Next") & "</FONT>"
				End If

				strHTML = strHTML & "  </TD>"
				strHTML = strHTML & "</TR>"

			Else
						
				' Do nothing
							
			End If
			
			Response.Write strHTML
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

</P>
</BODY>
</HTML>

<%
	cnnDB.Close
	Set cnnDB = Nothing
%>