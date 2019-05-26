<% 
Option Explicit

Response.Buffer = True	'Buffer the response, so Response.Expires can be used
%>
<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: rptCategorySummary.asp
'  Date:     $Date: 2004/03/24 23:53:14 $
'  Version:  $Revision: 1.6 $
'  Purpose:  Produces a SCategory Summary report for all cases
' ----------------------------------------------------------------------------------
%>

<!-- #Include File = "Include/Public.asp" -->

<html>

<head>
	<title>Reports</title>
</head>

<link rel="stylesheet" type="text/css" href="Default.css">

<%

Dim cnnDB
Dim intUserID
Dim strSQL, strTotals
Dim rstResults, rstTotals
Dim dtStartDate, dtEndDate
Dim binUserPermMask
Dim strHTML, strWhere


Set cnnDB = CreateConnection
	    
intUserID = GetUserID
binUserPermMask = GetUserPermMask


' Check permissions

If PERM_ACCESS_REPORTS = (PERM_ACCESS_REPORTS And binUserPermMask) Then
	' Report access granted
Else
	DisplayError 4, ""
End If


' Get date constraints
strWhere = ""
If Not IsEmpty(Request.Form("cbxStartMonth")) Then
	dtStartDate = Request.Form("cbxStartMonth") & "-" & Request.Form("cbxStartDay") & "-" & Request.Form("cbxStartYear") & " 00:00:00"
	dtEndDate = Request.Form("cbxEndMonth") & "-" & Request.Form("cbxEndDay") & "-" & Request.Form("cbxEndYear") & " 23:59:59"
	strWhere = "WHERE tblCases.RaisedDate>'" & dtStartDate & "' AND tblCases.RaisedDate<'" & dtEndDate & "' "
End If

strSQL =	"SELECT tblCategories.CatName, Count(tblCases.CasePK) AS NoOfCases " & _
			", ISNull(SUM(tblNotes.MinutesSpent),0) AS MinutesSpent , " & _
			"ISNull(SUM(tblNotes.MinutesSpent),0)/Count(tblCases.CasePK) AS AvgMinutesSpent " & _
			"FROM tblCases  " & _
			"INNER JOIN tblCategories ON tblCases.CatFK = tblCategories.CatPK  " & _
			"LEFT OUTER JOIN  " & _
					"(SELECT tblNotes.CaseFK, ISNull(SUM(tblNotes.MinutesSpent),0) as MinutesSpent  " & _
					"FROM tblNotes GROUP BY tblNotes.CaseFK)  " & _
				"as tblNotes ON tblCases.CasePK = tblNotes.CaseFK " & _
			strWhere & _
			"GROUP BY tblCategories.CatName  " 
		
Set rstResults = Server.CreateObject("ADODB.Recordset")
	
rstResults.Open strSQL, cnnDB

strTotals =	"SELECT Count(tblCases.CasePK) AS NoOfCases " & _
			", ISNull(SUM(tblNotes.MinutesSpent),0) AS MinutesSpent , " & _
			"ISNull(SUM(tblNotes.MinutesSpent),0)/Count(tblCases.CasePK) AS AvgMinutesSpent " & _
			"FROM tblCases  " & _
			"INNER JOIN tblCategories ON tblCases.CatFK = tblCategories.CatPK  " & _
			"LEFT OUTER JOIN  " & _
					"(SELECT tblNotes.CaseFK, ISNull(SUM(tblNotes.MinutesSpent),0) as MinutesSpent  " & _
					"FROM tblNotes GROUP BY tblNotes.CaseFK)  " & _
				"as tblNotes ON tblCases.CasePK = tblNotes.CaseFK " & _
			strWhere 

Set rstTotals = Server.CreateObject("ADODB.Recordset")
	
rstTotals.Open strTotals, cnnDB
	
If (rstResults.BOF And rstResults.EOF) OR (rstTotals.BOF And rstTotals.EOF) Then
	
	' No records returned
	
Else
	strHTML = ""
	While Not(rstResults.EOF) 
		
		strHTML = strHTML & "<TR style=""FONT-SIZE: 9pt"">" & Chr(13)
		strHTML = strHTML & "  <TD align=Left>&nbsp;" & rstResults.Fields("CatName") & "</TD>" & Chr(13)
		strHTML = strHTML & "  <TD align=Center>" & rstResults.Fields("NoOfCases") & "</TD>" & Chr(13)
		strHTML = strHTML & "  <TD align=Center>" & rstResults.Fields("MinutesSpent") & "</TD>" & Chr(13)
		strHTML = strHTML & "  <TD align=Center>" & rstResults.Fields("AvgMinutesSpent") & "</TD>" & Chr(13)
		strHTML = strHTML & "  <TD align=Center>" & _
							Round(CLng(rstResults.Fields("NoOfCases")) / CLng(rstTotals.Fields("NoOfCases")) * 100,1) & _
							"%</TD>" & Chr(13)
		strHTML = strHTML & "  <TD align=Center>" & _
							Round(CLng(rstResults.Fields("MinutesSpent")) / CLng(rstTotals.Fields("MinutesSpent")) * 100,1) & _
							"%</TD>" & Chr(13)
		strHTML = strHTML & "</TR>" & Chr(13)
						 				
		rstResults.MoveNext
			
	WEnd
	
		strHTML = strHTML & "<TR style=""FONT-SIZE: 9pt"">" & Chr(13)
		strHTML = strHTML & "  <TH align=Left>&nbsp;Totals</TH>" & Chr(13)
		strHTML = strHTML & "  <TH align=Center>" & rstTotals.Fields("NoOfCases") & "</TH>" & Chr(13)
		strHTML = strHTML & "  <TH align=Center>" & rstTotals.Fields("MinutesSpent") & "</TH>" & Chr(13)
		strHTML = strHTML & "  <TH align=Center>" & rstTotals.Fields("AvgMinutesSpent") & "</TH>" & Chr(13)
		strHTML = strHTML & "  <TH align=Center>100%</TH>" & Chr(13)
		strHTML = strHTML & "  <TH align=Center>100%</TH>" & Chr(13)
		strHTML = strHTML & "</TR>" & Chr(13)

	
End If
	
rstResults.Close
Set rstResults = Nothing
rstTotals.Close
Set rstTotals = Nothing
			 
%>

<p align=center>
<table class=Normal align=center cellSpacing=1 cellPadding=1 width="680" border=0>
  
  <tr>
    <td>
    <%
    Response.Write DisplayHeader
    %>
    </td>
  </tr>

  <tr>
    <td>
      <table class="lhd_Box" cellSpacing=0 cellPadding=1 width="100%" border=0 bgColor=white>
	    <tr class="lhd_Heading1">
		  <td colspan=6 align=center><%=Lang("Reports_Menu")%></td>
	    </tr>
	    <tr>
			<th width="15%" align="Left">
				&nbsp;<%=Lang("Category")%></th>
			<th width="15%" align="Center">
				<%=Lang("Case_Plural")%>
			</th>
			<th width="15%" align="Center">
				<%=Lang("Min_Spent")%>
			</th>
			<th width="15%" align="Center">
				<%=Lang("Avg_Min_Spent")%>
			</th>
			<th width="15%" align="Center">
				<%=Lang("Percentage_Of_Cases")%>
			</th>
			<th width="15%" align="Center">
				<%=Lang("Percentage_Of_Time")%>
			</th>
		</tr>
		<%
		Response.Write strHTML
		%>
		<tr>
	  </table>
	</td>
  </tr>

  <tr>
    <td>
    <%
    Response.Write DisplayFooter
    %>
    </td>
  </tr>
  
</table>
</p>

<body>

</html>

<%

cnnDB.Close
Set cnnDB = Nothing

%>
