<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: fileView.asp
'  Date:     $Date: 2004/03/11 06:08:42 $
'  Version:  $Revision: 1.2 $
'  Purpose:  This page provides the ability to view a stored attachment
' ----------------------------------------------------------------------------------
%>

<!-- #Include File = "Include/Public.asp" -->

<%
Dim oConn, oRs
Dim strSQL
Dim nFileID
Dim fsoFile


nFileID = Request.QueryString("id")

If Not nFileID = "" And IsNumeric(nFileID) Then

	Set oConn = CreateConnection
	Set oRs = Server.CreateObject("ADODB.Recordset")

	' Sometimes I personally have errors with one method on different servers, but the other works.
'	oConn.Open "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("Files.mdb")
	'oConn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("Files.mdb")

	strSQL = "SELECT FileName, ContentType, FileData, FileLocation FROM tblFiles WHERE FilePK = " & nFileID

	oRs.Open strSQL, oConn, 3, 3

	If Not oRs.EOF Then

    Set fsoFile = Server.CreateObject("Scripting.FileSystemObject")
    
    If fsoFile.FileExists(oRs.Fields("FileLocation").Value) Then
'      Response.Redirect "Attachments/" & Mid(oRs.Fields("FileLocation").Value, InstrRev(oRs.Fields("FileLocation").Value, "\")+1, Len(oRs.Fields("FileLocation").Value) - InstrRev(oRs.Fields("FileLocation").Value, "\"))
      Response.Redirect "/" & Mid(Request.ServerVariables("PATH_INFO"), 2, InStr(2, Request.ServerVariables("PATH_INFO"), "/")-2) & "/Attachments/" & Mid(oRs.Fields("FileLocation").Value, InstrRev(oRs.Fields("FileLocation").Value, "\")+1, Len(oRs.Fields("FileLocation").Value) - InstrRev(oRs.Fields("FileLocation").Value, "\"))
    Else
    	DisplayError 3, "Attachment not found, please contact your systems administrator."
    End If
    
    Set fsoFile = Nothing

	Else
	
  	DisplayError 3, "The attachment you are trying to view is not available or has been detached from this case, please refresh the page and try again or contact your systems administrator."
  	
	End If

	oRs.Close
	oConn.Close

	Set oRs = Nothing
	Set oConn = Nothing
	
Else

	DisplayError 3, "The attachment you are trying to view is not available or has been detached from this case, please refresh the page and try again or contact your systems administrator."
	
End If
%>

