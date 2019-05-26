<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: fileAttachment.asp
'  Date:     $Date: 2004/03/25 03:13:48 $
'  Version:  $Revision: 1.3 $
'  Purpose:  Provides the interface to allow a user to attach a file to a case
' ----------------------------------------------------------------------------------
%>

<!-- #Include File = "Include/Public.asp" -->
<!-- #Include File = "Include/Settings.asp" -->

<!-- #Include File = "Classes/clsUpload.asp" -->


<%
Dim cnnDB
Dim strCaseID
Dim intUserID
Dim binUserPermMask


Set cnnDB = CreateConnection

intUserID = GetUserID
binUserPermMask = GetUserPermMask

If Application("ENABLE_ATTACHMENTS") = 1 Then

  ' Do nothing, as attachments have been enabled

Else
  
  DisplayError 3, "Attachments have been disabled, please contact your system administrator for assistance."

End If


' Note that the CaseID is passed as a string.  This is because the field in the 
' database is a text field as we want to be able to create attachments without
' having an assign CaseID

strCaseID = Request.QueryString("id")

%>

<HTML>

<SCRIPT for="btnClose" event="onClick" language="VBScript">

	window.close 
	
</SCRIPT>

<SCRIPT for="btnView" event="onClick" language="VBScript">

	window.open "fileView.asp?ID=" & document.frmFiles.lbxFiles.value
	
</SCRIPT>


<LINK rel="stylesheet" type="text/css" href="Default.css">

<P align="Center">

<TABLE class="Normal" style="WIDTH: 450px" cellSpacing="0" cellPadding="0">
	<TR>
		<TD>
		</TD>
	</TR>
	<TR>
		<TD align="center">
			<TABLE class="lhd_Box" style="WIDTH: 100%" cellSpacing="0" cellPadding="0">
				<FORM method="post" encType="multipart/form-data" action="fileUpload.asp?Mode=1&CaseID=<%=strCaseID%>&UserID=<%=intUserID%>" id=frmAttach name=frmAttach>
				<TR class="lhd_Heading1">
					<TD colspan=4><%=Lang("Attach_Files")%></TD>
				</TR>
				<TR>
					<TD width=5%>
					</TD>
					<TD width=70%>
					</TD>
					<TD width=20%>
					</TD>
					<TD width=5%>
					</TD>
				</TR>
				<TR>
					<TD>
					</TD>
					<TD align=left colspan=2>
						1. Click <B>Browse</B> to select the file, or type the path to the file in the box below.
					</TD>
					<TD>
					</TD>
				</TR>
				<TR>
					<TD>
					</TD>
					<TD align=right colspan=2>
						<INPUT style="WIDTH: 100%;" type="File" name="File1">
					</TD>
					<TD>
					</TD>
				</TR>
				<TR>
					<TD colspan=4>
					</TD>
				</TR>
				<TR>
					<TD>
					</TD>
					<TD align=left colspan=2>
						2. Attach the file to the Case by clicking <B>Attach</B>.  File Transfer times vary depending upon file size.
					</TD>
					<TD>
					</TD>
				</TR>
				<TR>
					<TD>
					</TD>
					<TD align=right colspan=2>
						<INPUT style="WIDTH: 80px" type="Submit" value="<%=Lang("Attach")%>" id=btnAttach name=btnAttach>
					</TD>
					<TD>
					</TD>
				</TR>
				</FORM>
				<TR>
					<TD colspan=4>
					</TD>
				</TR>
				<TR>
					<TD>
					</TD>
					<TD align=left colspan=3>
						3. View the list of files attached to this case.  (To remove an attached file
						   select and highlight the filename and click <B>Remove</B>).
					</TD>
				</TR>
				<FORM method="POST" action="fileUpload.asp?Mode=2&CaseID=<%=strCaseID%>&UserID=<%=intUserID%>" id=frmFiles name=frmFiles>
				<TR>
					<TD>
					</TD>
					<TD valign=top align=left>
						<SELECT size=5 style="WIDTH: 100%" id=lbxFiles name=lbxFiles>
							<%
							Set rsAttachments = Server.CreateObject("ADODB.Recordset")
							
							rsAttachments.Open "SELECT * FROM tblFiles WHERE CaseFK='" & strCaseID & "'", cnnDB
							
							If rsAttachments.BOF And rsAttachments.EOF Then
							
								' No records found
								
							Else
							
								Do While Not rsAttachments.EOF
								%>
									<OPTION value="<%=rsAttachments.Fields("FilePK")%>"><%=rsAttachments.Fields("FileName")%> (<%=Int(rsAttachments.Fields("FileSize")/1024)%>k)</OPTION>
								<%
									rsAttachments.MoveNext
								Loop
							
							End If
							
							rsAttachments.Close
							Set rsAttachments = Nothing
							
							%>
						</SELECT>
					</TD>
					<TD valign=bottom align=right>
						<INPUT style="WIDTH: 80px" type="Button" value="<%=Lang("View")%>" id=btnView name=btnView>
						<INPUT style="WIDTH: 80px" type="Submit" value="<%=Lang("Remove")%>" id=btnRemove name=btnRemove>
					</TD>
					<TD>
					</TD>
				</TR>
				</FORM>
				<TR>
					<TD colspan=4>
					</TD>
				</TR>
				<TR>
					<TD>
					</TD>
					<TD align=center colspan=2>
						<INPUT style="WIDTH: 110px; BACKGROUND-COLOR: white" type="button" value="<%=Lang("Close")%>" id=btnClose name=btnClose>
					</TD>
					<TD>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD>
		</TD>
	</TR>
</TABLE>

</P>

</HTML>

<%

cnnDB.Close
Set cnnDB = Nothing

%>