<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: rptMenu.asp
'  Date:     $Date: 2004/03/11 06:32:56 $
'  Version:  $Revision: 1.5 $
'  Purpose:  Provides the reporting main menu
' ----------------------------------------------------------------------------------
%>

<!--METADATA TYPE="TypeLib" NAME="Microsoft Scripting Runtime" UUID="{420B2830-E718-11CF-893D-00A0C9054228}" VERSION="1.0"-->

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

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<SCRIPT language="JavaScript">

	var calbind				//binds to the object the calendar wants to change
	var cwindowsettings		//settings for the calender windows

	//this function is called by the calender popup so that the date boxes can be changed to he date selected by the user
	function changeItem(objstring){ //when the calender has been cklicked the date is changed to the item specified in tagname
		calbind.value = objstring
	}

	//Is called to produce a calendar
	function makeCalendar(targetobj, oDate){ //this funtion creates the calender

		cwindowsettings = "fullscreen=no, toolbar=no, status=no, menubar=no, scrollbars=no, resizable=no, directories=no, location=no, "
		cwindowsettings = cwindowsettings + "left=" + (window.event.x)+ ", top=" + (window.event.y+95) + ", width=196, height=185"

		calbind = eval(targetobj)

		window.open("objCalendar.asp", "CalendarWindow", cwindowsettings, true)

	}
	
  // sets the form action to open the chosen report
	function SubmitForm() { 


    if (document.frmReports.lbxReports.options(document.all("lbxReports").selectedIndex).value != "")
    {

    }
    else
    {
    
    }

		switch(document.frmReports.lbxReports.options(document.all("lbxReports").selectedIndex).value) 
		{
			case "rpt1":
				document.frmReports.action = "rptCategorySummary.asp"
				break

			case "rpt2":
				document.frmReports.action = "rptStatusSummary.asp"
				break

			default :
				break
		}
		
		document.frmReports.submit  // submit the form
		    
  }	

</SCRIPT>
<HTML>
	<HEAD>
		<META content="MSHTML 6.00.2600.0" name="GENERATOR">
	</HEAD>
	<%
	Dim cnnDB
	Dim intUserID
	Dim blnUserPermMask, binRequiredPerm


	' Get user variables

	Set cnnDB = CreateConnection

	intUserID = GetUserID
	binUserPermMask = GetUserPermMask


	' Check permissions

	If PERM_ACCESS_REPORTS = (PERM_ACCESS_REPORTS And binUserPermMask) Then
		' Report access granted
	Else
		DisplayError 4, ""
	End If

%>
	<LINK rel="stylesheet" type="text/css" href="Default.css">
		<BODY>
			<P align="center">
				<TABLE class="Normal" align="center" cellSpacing="1" cellPadding="1" width="680" border="0">
					<TR>
						<TD>
							<%
						Response.Write DisplayHeader
						%>
						</TD>
					</TR>
					<TR>
						<TD>
							<TABLE class="lhd_Box" cellSpacing="0" cellPadding="1" width="100%" border="0" bgColor="white">
								<FORM name="frmReports" action="" method="POST">
									<TR class="lhd_Heading1">
										<TD colspan="5" align="center"><%=Lang("Reports_Menu")%></TD>
									</TR>
									<TR>
										<TD width="20%"></TD>
										<TD width="20%"></TD>
										<TD width="20%"></TD>
										<TD width="20%"></TD>
										<TD width="20%"></TD>
									</TR>
									<TR class="lhd_Heading2">
										<TD></TD>
										<TD colspan="3"><%=Lang("Reports")%></TD>
										<TD></TD>
									</TR>
									<TR>
										<TD></TD>
										<TD colspan="3">
											<SELECT style="WIDTH: 100%;" size="7" id="lbxReports" name="lbxReports">
												<OPTION value="rpt1" selected>Category Summary Report</OPTION>
												<OPTION value="rpt2">Status Summary Report</OPTION>
											</SELECT>
										</TD>
										<TD></TD>
									</TR>
									<TR>
										<TD></TD>
										<TD colspan="3"></TD>
										<TD></TD>
									</TR>
									<TR class="lhd_Heading2">
										<TD></TD>
										<TD colspan="3"><%=Lang("Report_Criteria")%></TD>
										<TD></TD>
									</TR>
									<TR>
										<TD></TD>
										<TD colspan="3">
											<TABLE width="100%" border="0" cellpadding="0" cellspacing="0">
												<TR>
													<TD width="50%"><INPUT type="radio" id="radAllAvailable" name="radio2" checked>&nbsp;All&nbsp;&nbsp;&nbsp;<INPUT type="radio" id="radCustom" name="radio2">&nbsp;Custom 
														Range</TD>
													<TD width="50%" align="Right">
														Date From:&nbsp; <INPUT style="WIDTH=95px;" type="text" id="tbxDateFrom" name="tbxDateFrom">
														<A href="javascript: ;" OnClick="javascript:makeCalendar('document.frmReports.tbxDateFrom', 'document.frmReports.tbxDateFrom.value') ;">
															<IMG src="Images/Calendar.gif" border="0" ALIGN="absmiddle"></A>
													</TD>
												</TR>
												<TR>
													<TD></TD>
													<TD align="Right">
														Date To:&nbsp;<INPUT style="WIDTH=95px;" type="text" id="tbxDateTo" name="tbxDateTo">
														<A href="javascript: ;" OnClick="javascript:makeCalendar('document.frmReports.tbxDateTo', 'document.frmReports.tbxDateTo.value') ;">
															<IMG src="Images/Calendar.gif" border="0" ALIGN="absmiddle"></A>
													</TD>
												</TR>
											</TABLE>
										</TD>
										<TD></TD>
									</TR>
									<TR>
										<TD colspan="5"></TD>
									</TR>
									<TR>
										<TD></TD>
										<TD colspan="3" align="center">
											<INPUT type="submit" value="Generate Report" name="btnGenerate" id="btnGenerate" style="WIDTH: 120px; BACKGROUND-COLOR: white"
												onclick="Javascript:SubmitForm()">
										</TD>
										<TD></TD>
									</TR>
									<TR>
										<TD colspan="5"></TD>
									</TR>
								</FORM>
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
