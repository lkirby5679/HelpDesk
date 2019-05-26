<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: Menu.asp
'  Date:     $Date: 2004/03/30 04:20:30 $
'  Version:  $Revision: 1.7 $
'  Purpose:  The main menu page
' ----------------------------------------------------------------------------------
%>

<%
Option Explicit
%>


<!-- #Include File = "Include/Settings.asp" -->
<!-- #Include File = "Include/Public.asp" -->


<%

Dim cnnDB
Dim intUserID
Dim binUserPermMask, binRequiredPerm



' Create the connection to the database
Set cnnDB = CreateConnection

' Determine the logged in User's ID
intUserID = GetUserID
binUserPermMask = GetUserPermMask


%>

<HEAD>
	
	<META content="Microsoft FrontPage 6.0" name=GENERATOR>
</HEAD>

<LINK rel="stylesheet" type="text/css" href="Default.css">

<BODY>
<P align=center>
<table class="Normal">
	<tr>
		<td>
		  <% Response.Write DisplayHeader() %>
			<a href="http://www.transworldinteractive.net/">
			<img border="0" src="Images/TIisIT.jpg" width="407" height="104"></a></td>
	</tr>
	<tr>
		<td>
			<table class="lhd_Box" cellSpacing="0">
				<tr class="lhd_Heading1">
					<td colspan=2 align=middle><%=Lang("Main_Menu")%></td>
				</tr>
			  <tr>
			    <td width="50%" valign="Top">
			      <table width="100%">
				      <tr>
				      	<TD width="5%">&nbsp;</td>
				      	<TD width="95%">&nbsp;</td>
				      </tr>
				      <tr class="lhd_Heading2">
				      	<td></td>
				      	<td>Case Options</td>
				      </tr>
				      <tr>
				      	<td></td>
				      	<td valign="Top">
				          <b>
				      	  <blockquote>
				      	    <a href="caseNew.asp"><img border="0" src="Images/New.gif">&nbsp;<%=Lang("Submit_New_Case")%></a>
				      	    <br>
				      	    <br>
				      	    <a href="caseList.asp?Mode=1&Column=1&Order=0&Page=1"><img border="0" src="Images/List.gif">&nbsp;<%=Lang("List_My_Active_Cases")%></a>
				      	    <br>
				      	    <br>
				      	    <a href="caseSearch.asp"><img border="0" src="Images/Search.gif">&nbsp;<%=Lang("Search_Cases")%></a>
				      	  </blockquote>
				          </b>
				      	</td>
				      </tr>
				      <%
				      If Application("ENABLE_KB") = 1 Then

				        If (PERM_KB_READ = (PERM_KB_READ And binUserPermMask)) Or (PERM_KB_MODIFY = (PERM_KB_MODIFY And binUserPermMask)) Or (PERM_KB_CREATE = (PERM_KB_CREATE And binUserPermMask))Then
                %>
				          <tr>
				            <td colspan="2"></td>
				          </tr>
				          <tr class="lhd_Heading2">
				          	<td></td>
				          	<td>Knowledgebase</td>
				          	<td></td>
				          	<td></td>
				          </tr>
				          <%
  				        If PERM_KB_CREATE = (PERM_KB_CREATE And binUserPermMask) Then
  				        %>
				            <tr>
				            	<td></td>
				            	<td valign="Top">
				                <b>
				            	  <blockquote>
				            	    <a href="kbRecord.asp"><img border="0" src="Images/New.gif">&nbsp;<%=Lang("Add_to_Knowledgebase")%></a>
				            	  </blockquote>
				                </b>
				              </td>
				            	<td></td>
				            	<td></td>
				            </tr>
				          <%
				          Else
				            ' Do nothing
				          
				          End If
				          %>
				          <tr>
				          	<td></td>
				          	<td valign="Top">
				              <b>
				          	  <blockquote>
				          	    <a href="kbSearch.asp"><img border="0" src="Images/Search.gif">&nbsp;<%=Lang("Search_Knowledgebase")%></a>
				          	  </blockquote>
				              </b>
				            </td>
				          	<td></td>
				          	<td></td>
				          </tr>
				        <%
				        Else
				      
				          ' Do nothing
				          
				        End If
				        
				      Else
				      
				        ' Do nothing
				        
				      End If
				      %>
				      <%
				      If Application("ENABLE_REPORTS") = 1 Then

				        If PERM_ACCESS_REPORTS = (PERM_ACCESS_REPORTS And binUserPermMask) Then
                %>
				          <tr>
				            <td colspan="2"></td>
				          </tr>
				          <tr class="lhd_Heading2">
				          	<td></td>
				          	<td>Reporting</td>
				          	<td></td>
				          	<td></td>
				          </tr>
				          <tr>
				          	<td></td>
				          	<td valign="Top">
				              <b>
				          	  <blockquote>
				          	    <a href="rptMenu.asp"><img border="0" src="Images/Report.gif">&nbsp;<%=Lang("View_Reports")%></a>
				          	  </blockquote>
				              </b>
				            </td>
				          	<td></td>
				          	<td></td>
				          </tr>
				        <%
				        Else
				          
				          ' Do nothing
				            
				        End If
				        
				      Else
				      
				        ' Do nothing
				      
				      End If
				      %>
				    </table>
				  </td>
				  <td width="50%" valign="Top">
				    <table width="100%">
				      <tr>
				      	<TD width="5%">&nbsp;</td>
				      	<TD width="95%">&nbsp;</td>
				      </tr>
				      <tr class="lhd_Heading2">
				      	<td></td>
				      	<td>Profile Options</td>
				      </tr>
				      <tr>
				      	<td></td>
				      	<td valign="Top">
				      	  <b>
				      	  <blockquote>
				      	    <a href="Contact.asp?ID=<%=intUserID%>"><img border="0" src="Images/Tools.gif">&nbsp;<%=Lang("Edit_My_Profile")%></a>
				      	  </blockquote>
				      	  </b>
				      	</td>
				      </tr>
				      <%
				      If PERM_ACCESS_TECH = (PERM_ACCESS_TECH And binUserPermMask) Then
              %>
				        <tr>
				          <td colspan="2"></td>
				        </tr>
				        <tr class="lhd_Heading2">
				        	<td></td>
				        	<td>Technician Options</td>
				        </tr>
				        <tr>
				        	<td></td>
				        	<td valign="Top">
				            <b>
				        	  <blockquote>
				        	    <a href="caseList.asp?Mode=3&Column=1&Order=0&Page=1"><img border="0" src="Images/List.gif">&nbsp;<%=Lang("List_My_Assigned_Cases")%></a>
				        	    <br>
				        	    <br>
				        	    <a href="caseList.asp?Mode=2&Column=1&Order=0&Page=1"><img border="0" src="Images/List.gif">&nbsp;<%=Lang("List_All_Un-Assigned_Cases")%></a>
				        	  </blockquote>
				            </b>
				          </td>
				        </tr>
				      <%
				      Else
				      
				        ' Do nothing
				      
				      End If
				      %>
				      
				      <%
 				      If Application("ENABLE_INOUT") = 1 Then
              %>
				        <tr>
				          <td colspan="2"></td>
				        </tr>
      					<tr class="lhd_Heading2">
				          	<td></td>
				          	<td>In/Out Board</td>
				        </tr>
				        <tr>
				          	<td></td>
				          	<td valign="Top">
				                <b><blockquote>
				          	    <a href="inoutList.asp"><img border="0" src="Images/List.gif">&nbsp;In/Out Status</a>
				          	    <br><br>
						            <a href="inoutPhoneList.asp"><img border="0" src="Images/List.gif">&nbsp;Phone list</a>
				          	</blockquote></b>
				            </td>
				        </tr>
				      <%
				      Else
				      
				        ' Do nothing
				        
				      End If
				      %>
				      
				      <%
				      If PERM_ACCESS_ADMIN = (PERM_ACCESS_ADMIN And binUserPermMask) Then
              %>
				        <tr>
				          <td colspan="2"></td>
				        </tr>
				        <tr class="lhd_Heading2">
				        	<td></td>
				        	<td>Administration</td>
				        </tr>
				        <tr>
				        	<td></td>
				        	<td valign="Top">
				            <b>
				        	  <blockquote>
				        	    <a href="admMenu.asp"><img border="0" src="Images/Tools.gif">&nbsp;<%=Lang("Manage_&_Configure_System")%></a>
				        	  </blockquote>
				            </b>
				          </td>
				        </tr>
				      <%
				      Else
				      
				        ' Do nothing
				        
				      End If
				      %>
				    </table>
				  </td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		  <% Response.Write DisplayFooter() %>
		</td>
	</tr>
</table>

</P>
</BODY>

</HTML>

<%
cnnDB.Close
Set cnnDB = Nothing
%>