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
'  Filename: rptStatusSummary.asp
'  Date:     $Date: 2004/03/11 06:32:56 $
'  Version:  $Revision: 1.5 $
'  Purpose:  Produces a Status Summary report for all cases
' ----------------------------------------------------------------------------------
%>


<!-- #Include File = "Include/Public.asp" -->
<%
Dim cnnDB
Dim intUserID
Dim strSQL, strHTML
Dim rstResults
Dim dtStartDate, dtEndDate
Dim binUserPermMask


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
	
dtStartDate = Request.Form("cbxStartMonth") & "/" & Request.Form("cbxStartDay") & "/" & Request.Form("cbxStartYear") & " 00:00:00"
dtEndDate = Request.Form("cbxEndMonth") & "/" & Request.Form("cbxEndDay") & "/" & Request.Form("cbxEndYear") & " 23:59:59"

strSQL = strSQL & "AND tblCases.RaisedDate>" & dtStartDate & " AND tblCases.RaisedDate<" & dtEndDate & " "

' Build the query string

strSQL = ""
strSQL = strSQL & "SELECT * "
strSQL = strSQL & "FROM "
strSQL = strSQL & "(( "
strSQL = strSQL & "SELECT qryStatusTotals.RepFK AS RepFK, Max(OpenCases) AS NoOpenCases, Max(ClosedCases) AS NoClosedCases, Max(PendingCases) AS NoPendingCases, Max(CancelledCases) AS NoCancelledCases "
strSQL = strSQL & "FROM "
strSQL = strSQL & "( "
strSQL = strSQL & "SELECT tblCases.RepFK, Count(tblCases.StatusFK) As OpenCases, Null As ClosedCases, Null As PendingCases, Null As CancelledCases "
strSQL = strSQL & "FROM tblCases "
strSQL = strSQL & "WHERE tblCases.StatusFK=(SELECT tblLists.ListItemPK FROM tblLists WHERE tblLists.ItemName='Open') "
strSQL = strSQL & "GROUP BY tblCases.StatusFK, tblCases.RepFK "

strSQL = strSQL & "UNION ALL "
	
strSQL = strSQL & "SELECT tblCases.RepFK, Null As OpenCases, Null As ClosedCases, Count(tblCases.StatusFK) As PendingCases, Null As CancelledCases "
strSQL = strSQL & "FROM tblCases "
strSQL = strSQL & "WHERE tblCases.StatusFK=(SELECT tblLists.ListItemPK FROM tblLists WHERE tblLists.ItemName='Pending') "
strSQL = strSQL & "GROUP BY tblCases.StatusFK, tblCases.RepFK "
	
strSQL = strSQL & "UNION ALL "
	
strSQL = strSQL & "SELECT tblCases.RepFK, Null As OpenCases, Null As ClosedCases, Null As PendingCases, Count(tblCases.StatusFK) As CancelledCases "
strSQL = strSQL & "FROM tblCases "
strSQL = strSQL & "WHERE tblCases.StatusFK=(SELECT tblLists.ListItemPK FROM tblLists WHERE tblLists.ItemName='Cancelled') "
strSQL = strSQL & "GROUP BY tblCases.StatusFK, tblCases.RepFK "

strSQL = strSQL & "UNION ALL "
	
strSQL = strSQL & "SELECT tblCases.RepFK, Null As OpenCases, Count(tblCases.StatusFK) As ClosedCases, Null As PendingCases, Null As CancelledCases "
strSQL = strSQL & "FROM tblCases "
strSQL = strSQL & "WHERE tblCases.StatusFK=(SELECT  tblLists.ListItemPK FROM tblLists WHERE tblLists.ItemName='Closed') "
strSQL = strSQL & "GROUP BY tblCases.StatusFK, tblCases.RepFK "
strSQL = strSQL & ") AS qryStatusTotals "
strSQL = strSQL & "GROUP BY qryStatusTotals.RepFK "
strSQL = strSQL & ") AS qryStatuses "
	
strSQL = strSQL & "RIGHT JOIN "

strSQL = strSQL & "( "
strSQL = strSQL & "SELECT tblContacts.UserName AS UserName, tblContacts.FName AS FName, tblContacts.LName AS LName, tblGroupMembers.ContactFK AS ContactFK "
strSQL = strSQL & "FROM tblContacts INNER JOIN tblGroupMembers ON tblContacts.ContactPK = tblGroupMembers.ContactFK "
strSQL = strSQL & ") AS qryReps "

strSQL = strSQL & "ON qryReps.ContactFK = qryStatuses.RepFK "
strSQL = strSQL & ") "
strSQL = strSQL & "ORDER BY UserName ASC "

			 
Set rstResults = Server.CreateObject("ADODB.Recordset")
	
rstResults.Open strSQL, cnnDB
	
If rstResults.BOF And rstResults.EOF Then
	
	' No records returned
	
Else
	
	While Not(rstResults.EOF) 
		
    strHTML = strHTML & "<TR style=""FONT-SIZE: 9pt"">" & Chr(13)
    strHTML = strHTML & "  <TD align=Left>&nbsp;" & rstResults.Fields("UserName") & "&nbsp;&nbsp;(&nbsp;" & rstResults.Fields("FName") & " " & rstResults.Fields("LName") & "&nbsp;)</TD>" & Chr(13)

    If Len(rstResults.Fields("NoOpenCases")) > 0 Then
    	strHTML = strHTML & "  <TD align=Center>" & rstResults.Fields("NoOpenCases") & "</TD>" & Chr(13)
    Else
    	strHTML = strHTML & "  <TD align=Center>-</TD>" & Chr(13)
    End If

    If Len(rstResults.Fields("NoPendingCases")) > 0 Then
    	strHTML = strHTML & "  <TD align=Center>" & rstResults.Fields("NoPendingCases") & "</TD>" & Chr(13)
    Else
    	strHTML = strHTML & "  <TD align=Center>-</TD>" & Chr(13)
    End If

    If Len(rstResults.Fields("NoCancelledCases")) > 0 Then
    	strHTML = strHTML & "  <TD align=Center>" & rstResults.Fields("NoCancelledCases") & "</TD>" & Chr(13)
    Else
    	strHTML = strHTML & "  <TD align=Center>-</TD>" & Chr(13)
    End If

    If Len(rstResults.Fields("NoClosedCases")) > 0 Then
    	strHTML = strHTML & "  <TD align=Center>" & rstResults.Fields("NoClosedCases") & "</TD>" & Chr(13)
    Else
    	strHTML = strHTML & "  <TD align=Center>-</TD>" & Chr(13)
    End If

    strHTML = strHTML & "</TR>" & Chr(13)
    				 				
    rstResults.MoveNext
		
	WEnd
	
End If
	
rstResults.Close
Set rstResults = Nothing

%>
<html>
	<link rel="stylesheet" type="text/css" href="Default.css">
		<head>
			<title>
				<%=Lang("Reports")%>
			</title>
		</head>
		<p align="center">
			<table class="Normal" align="center" cellSpacing="1" cellPadding="1" width="680" border="0">
				<tr>
					<td>
						<%
    Response.Write DisplayHeader
    %>
					</td>
				</tr>
				<tr>
					<td>
						<table class="lhd_Box" cellSpacing="0" cellPadding="1" width="100%" border="0" bgColor="white">
							<tr class="lhd_Heading1">
								<td colspan="5" align="center"><%=Lang("Status_Summary_Report")%></td>
							</tr>
							<tr>
								<th width="40%" align="Left">
									&nbsp;<%=Lang("Rep")%></th>
								<th width="15%" align="Center">
									<%=Lang("Open")%>
								</th>
								<th width="15%" align="Center">
									<%=Lang("Pending")%>
								</th>
								<th width="15%" align="Center">
									<%=Lang("Cancelled")%>
								</th>
								<th width="15%" align="Center">
									<%=Lang("Closed")%>
								</th>
							</tr>
							<%
							Response.Write strHTML
							%>
							<tr>
							</tr>
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
