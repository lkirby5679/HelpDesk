<%
Option Explicit
%>
<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: fileUpload.asp
'  Date:     $Date: 2004/03/29 01:07:29 $
'  Version:  $Revision: 1.4 $
'  Purpose:  Used to upload/save the attachment to either a database or file system
' ----------------------------------------------------------------------------------
%>

<!-- #Include File = "Include/Public.asp"-->

<!-- #include file = "Classes/clsUpload.asp"-->

<%
Dim objUpload, objConn, objRs
Dim strFileName, strFileLocation
Dim lngFileID, lngCaseID
Dim intUserID, intMode
Dim blnSave, blnError
Dim dtUploadDate
Dim fsoFile


Set objConn = CreateConnection

intMode = Request.QueryString("mode")
intUserID = Request.QueryString("userid")
lngCaseID = Request.QueryString("caseid")

blnError = False

Select Case intMode

	Case 1	' Upload the specified file
	
		' Instantiate Upload Class
		Set objUpload = New clsUpload

		' Grab the file name
		dtUploadDate = SQLDateFormat(Day(Now()), Month(Now()), Year(Now())) & " " & FormatDateTime(Now(), vbShortTime)

		If objUpload.Fields("File1").Length > (Application("MAX_ATTACHMENT_SIZE") * 1000)Then

      blnError = True

		Else

		  Set objRs = Server.CreateObject("ADODB.Recordset")

		  objRs.Open "tblFiles", objConn, 3, 3

		  objRs.AddNew

		  objRs.Fields("CaseFK").Value = lngCaseID
		  objRs.Fields("FileName").Value = objUpload.Fields("File1").FileName
		  objRs.Fields("FileSize").Value = objUpload.Fields("File1").Length
		  objRs.Fields("ContentType").Value = objUpload.Fields("File1").ContentType

      ' Compile path to save file to
      If InStr(1, CStr(lngCaseID),"x") > 0 Then
        strFileName = "x" & CStr(lngCaseID) & "x" & objUpload.Fields("File1").FileName
      Else
        strFileName = CStr(lngCaseID) & "x" & Right("20" & Year(Now()), 4) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2) & Right("0" & Hour(Now()), 2) & Right("0" & Minute(Now()), 2) & Right("0" & Second(Now()), 2) & "x" & objUpload.Fields("File1").FileName
      End If
      
      strFileLocation = Server.MapPath("\" & Mid(Request.ServerVariables("PATH_INFO"), 2, InStr(2, Request.ServerVariables("PATH_INFO"), "/")-2) & "\Attachments") & "\" & strFileName

      ' Update the database record
      objRs.Fields("FileLocation").Value = strFileLocation

      ' Save the binary data to the file system
      objUpload("File1").SaveAs strFileLocation
      

		  objRs.Fields("UploadDate").Value = dtUploadDate
		  objRs.Fields("LastUpdate").Value = dtUploadDate
		  objRs.Fields("LastUpdateByFK").Value = intUserID

		  objRs.Update

  '		lngFileID = objRs.Fields("FilePK").Value 

		  objRs.Close
		  Set objRs = Nothing

		End If
			
		Set objUpload = Nothing
	

	Case 2	' Delete the select file
	
		lngFileID = Request.Form("lbxFiles")

		If Len(lngFileID) > 0 And IsNumeric(lngFileID) Then
		
		  Set objRs = Server.CreateObject("ADODB.Recordset")
		  objRs.Open "SELECT * FROM tblFiles WHERE FilePK=" & lngFileID, objConn, 3, 3
	
      If objRs.BOF And objRs.EOF Then
      
        ' Record does not exist
        
      Else

		    ' Delete the database record
		  
			  objConn.Execute "DELETE FROM tblFiles WHERE FilePK=" & lngFileID

        ' Delete the physical file

		    Set fsoFile = Server.CreateObject("Scripting.FileSystemObject")

		    If fsoFile.FileExists(objRs.Fields("FileLocation").value) Then
		      fsoFile.DeleteFile objRs.Fields("FileLocation").value
		      ' File doesn't exist
		    End If
		
		    Set fsoFile = Nothing
		    
		  End If

		  objRs.Close
		  Set objRs = Nothing

		Else

			' Raise error, not a valid File ID

		End If
		
	
	Case Else
  	' Do nothing
		

End Select


objConn.Close
Set objConn = Nothing


If blnError = True Then

  Response.Write "<table width=""100%"">"
  Response.Write "  <tr>"
  Response.Write "    <td width=""10%""></td>"
  Response.Write "    <td width=""80%"">The attachment you are attempting to attach exceeds the max size allowed of " & Application("MAX_ATTACHMENT_SIZE") & "Kb.  Please contact your systems administrator if you have any queries.</td>"
  Response.Write "    <td width=""10%""></td>"
  Response.Write "  </tr>"
  Response.Write "  <tr>"
  Response.Write "    <td></td>"
  Response.Write "    <td></td>"
  Response.Write "    <td></td>"
  Response.Write "  </tr>"
  Response.Write "  <tr>"
  Response.Write "    <td></td>"
  Response.Write "    <td align=""Center""><a href=""javascript: void();"" onclick=""javascript: window.close()"">Close</a></td>"
  Response.Write "    <td></td>"
  Response.Write "  </tr>"
  Response.Write "</table>"
  
Else

  Response.Redirect "fileAttachment.asp?ID=" & lngCaseID
  
End If

%>
