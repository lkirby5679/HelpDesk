<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: caseSearch.asp
'  Date:     $Date: 2004/03/11 06:08:42 $
'  Version:  $Revision: 1.5 $
'  Purpose:  This page allows the user to search the case database
' ----------------------------------------------------------------------------------
%>
<%
Option Explicit
%>

<HTML>

<!--METADATA TYPE="TypeLib" NAME="Microsoft ActiveX Data Objects 2.5 Library" UUID="{00000205-0000-0010-8000-00AA006D2EA4}" VERSION="2.5"-->
<!--METADATA TYPE="TypeLib" NAME="Microsoft Scripting Runtime" UUID="{420B2830-E718-11CF-893D-00A0C9054228}" VERSION="1.0"-->
<!--METADATA TYPE="TypeLib" NAME="Microsoft CDO for Windows 2000 Library" UUID="{CD000000-8B95-11D1-82DB-00C04FB1625D}" VERSION="1.0"-->

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

<HEAD>


<META content="MSHTML 6.00.2600.0" name=GENERATOR></HEAD>

<LINK rel="stylesheet" type="text/css" href="Default.css">


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

</SCRIPT>

<%
	Dim cnnDB
	Dim binUserPermMask, binRequiredPerm
	Dim intUserID, intLastCaseType
	Dim strSQL, strHTML
	Dim rsR
	Dim objCollection


	' Get user variables

    Set cnnDB = CreateConnection
    
    intUserID = GetUserID
    binUserPermMask = GetUserPermMask


	' XML data island for CaseTypes and associated Categories

	strSQL = "SELECT tblCaseTypes.*, tblCategories.* FROM tblCaseTypes "
	strSQL = strSQL & "INNER JOIN tblCategories ON tblCaseTypes.CaseTypePK = tblCategories.CaseTypeFK "
	strSQL = strSQL & "WHERE tblCaseTypes.IsActive=" & lhd_True & " AND tblCategories.IsActive=" & lhd_True & ""
	
	Set rsR = server.createobject ("adodb.recordset")
	rsR.CursorLocation=adUseClient
	rsR.Open strSQL, cnnDB, adOpenStatic, adLockReadOnly, adCmdText
	
	strHTML = "<XML id=""XMLData1"">" & Chr(13) & "<CaseTypes>" & Chr(13)
	
	While Not rsR.EOF
	
		If rsR("CaseTypePK").Value <> intLastCaseType Then
			If Not IsEmpty(intLastCaseType) Then
				strHTML = strHTML & "	</CaseType>" & Chr(13)
			End If
			strHTML = strHTML & "	<CaseType CaseTypePK=" & Chr(34) & rsR("CaseTypePK").Value & Chr(34) & ">" & Chr(13)
			intLastCaseType = rsR("CaseTypePK").Value
		End If
   
		strHTML = strHTML & "		<Category>" & Chr(13)
		strHTML = strHTML & "			<CatPK>" & rsR("CatPK").value & "</CatPK>" & Chr(13)
		strHTML = strHTML & "			<CatName>" & rsR("CatName").value & "</CatName>" & Chr(13)
		strHTML = strHTML & "		</Category>" & Chr(13)
		rsR.MoveNext
   
	WEnd
	
	rsR.Close

	Set rsR = Nothing

	If Not IsEmpty(intLastCaseType)Then
		strHTML = strHTML & "	</CaseType>" & Chr(13)
	End If

	strHTML = strHTML & "</CaseTypes>" & Chr(13) & "</xml>" & Chr(13)
	
	Response.Write strHTML

%>	

<SCRIPT for="cbxCaseType" event="onChange" language="vbScript">
	
	' Generate list of associated Categories
	
	Set XML = Document.All("XMLData1")
	Set Nodes = XML.SelectNodes("CaseTypes/CaseType[@CaseTypePK='" & Document.All.cbxCaseType.Value & "']/Category")

	Set objCategoryList = Document.All("cbxCategory")

	objCategoryList.Options.Length = 1
	objCategoryList.Options(0).Value = 0
	objCategoryList.Options(0).InnerText = ""

	I = 1

	For Each Node In Nodes 
		strCatPK = Node.SelectSingleNode("CatPK").text
		strCatName = Node.SelectSingleNode("CatName").text
		objCategoryList.options.length = objCategoryList.options.length + 1 
		objCategoryList.options(I).Value = strCatPK
		objCategoryList.options(I).InnerText = strCatName
		
		I = I + 1
	Next

	Set objCategoryList = Nothing
	Set Nodes = Nothing
	Set XML = Nothing

</SCRIPT>

<BODY>

<P align="Center">

<TABLE width="680" cellSpacing="1">
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
	      <FORM action="caseSearchResults.asp?SQL=&Page=1" method="post" id=frmSearch name=frmSearch>
		    <TR class="lhd_Heading1">
		    	<TD colspan="7" align="Center"><%=Lang("Search")%></TD>
		    </TR>
		    <TR>
		      <TD width="8%"></TD>
		      <TD width="15%"></TD>
		      <TD width="27%"></TD>
		      <TD width="5%"></TD>
		      <TD width="15%"></TD>
		      <TD width="25%"></TD>
		      <TD width="5%"></TD>
		    </TR>
		    <TR class="lhd_Heading2">
		    	<TD colspan="7"><%=Lang("Keyword_Criteria")%></TD>
		    </TR>
		    <TR>
		      <TD></TD>
		      <TD><%=Lang("Keyword")%>:</TD>
		      <TD><INPUT type="text" style="WIDTH: 100%" id=tbxKeywords name=tbxKeywords></TD>
		      <TD></TD>
		      <TD></TD>
		      <TD></TD>
		      <TD></TD>
		    </TR>
		    <TR>
		      <TD colspan="7"></TD>
		    </TR>
		    <TR class="lhd_Heading2">
		    	<TD colspan="7"><%=Lang("Case_Criteria")%></TD>
		    </TR>
		    <TR>
		      <TD></TD>
		    	<TD><%=Lang("Contact")%>:</TD>
		    	<TD>
		    		<SELECT id=cbxContact style="WIDTH: 100%" name=cbxContact>
		    		<OPTION value="0" selected></OPTION>
		    		<%
		    		Set objCollection = New clsCollection
		    			
		    		objCollection.CollectionType = objCollection.clContact
		    		objCollection.Query = "SELECT ContactPK, UserName FROM tblContacts WHERE IsActive=" & lhd_True & " ORDER BY UserName ASC"
		    			
		    		If Not objCollection.Load Then
		    				
		    		    Response.Write objCollection.LastError
		    				    
		    		Else
		    					
		    			If objCollection.BOF And objCollection.EOF Then
		    					
		    				' No records
		    					
		    			Else
		    				
		    				Do While Not objCollection.EOF
		    				%>
		    					<OPTION VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.UserName%></OPTION>
		    				<%
		    					objCollection.MoveNext
		    				Loop
		    						
		    			End If
		    					
		    		End If
		    			
		    		Set objCollection = Nothing
		    		%>
		    		</SELECT>
		    	</TD>
		    	<TD></TD>
		      <TD></TD>
		    	<TD></TD>
		    	<TD></TD>
		    </TR>
		    <TR>
		      <TD></TD>
		    	<TD><%=Lang("Case_Type")%>:</TD>
		    	<TD>
		    		<SELECT id=cbxCaseType style="WIDTH: 100%" name=cbxCaseType>
		    	    <OPTION VALUE="0" SELECTED></OPTION>
		    	    <%
		    	    Set objCollection = New clsCollection
		    			    
		    	    objCollection.CollectionType = objCollection.clCaseType
		    	    objCollection.Query = "SELECT * FROM tblCaseTypes " &_
		    	                          "WHERE IsActive=" & lhd_True & " " &_
		    	                          "ORDER BY CaseTypePK ASC"
		    			    
		    	    If Not objCollection.Load Then
		    			    
		    			' Didn't load
		    					
		    		Else
		    			    
		    		    If objCollection.BOF And objCollection.EOF Then
		    			    
		    				' No records returned
		    						
		    			Else
		    					
		    				Do While Not objCollection.EOF
		    				%>
		    					<OPTION VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.CaseTypeName%></OPTION>
		    				<%
		    					objCollection.MoveNext
		    				Loop
		    					
		    			End If
		    			    
		    	    End If
		    	    %>
		    		</SELECT>
		    	</TD>
		    	<TD></TD>
	        <TD><%=Lang("Category")%>:</TD>
		    	<TD>
		    		<SELECT id=cbxCategory name=cbxCategory style="WIDTH: 100%" >
		    		<OPTION  value="0" selected></OPTION></SELECT>
		    	</TD>
		      <TD></TD>
		    </TR>
		    <TR>
		      <TD></TD>
		    	<TD><%=Lang("Priority")%>:</TD>
		    	<TD>
		    		<SELECT id=cbxPriority style="WIDTH: 100%" name=cbxPriority>
  	    		    <OPTION selected value="0"></OPTION>
		    		<%
		    		Response.Write BuildList("PRIORITY_LIST", 0)
		    	    %>
		    		</SELECT>
		    	</TD>
		      <TD></TD>
		      <TD></TD>
		    	<TD></TD>
		      <TD></TD>
		    </TR>
		    <TR>
		      <TD></TD>
 		    	<TD><%=Lang("Status")%>:</TD>
		    	<TD>
		    		<SELECT id=cbxStatus name=cbxStatus style="WIDTH: 100%">
  	    		    <OPTION selected value="0"></OPTION>
		    		<%
		    		Response.Write BuildList("STATUS_LIST", 0)
		    	    %>
		    		</SELECT>
		    	</TD>
		      <TD></TD>
		      <TD></TD>
		      <TD></TD>
		      <TD></TD>
		    </TR>
		    <TR>
		      <TD></TD>
		    	<TD><%=Lang("Assignment")%>:</TD>
		    	<TD>
		    		<SELECT id=cbxRep style="WIDTH: 100%" name=cbxRep>
		    		  <OPTION value="0" selected></OPTION>
		    		  <%
		    		  Set objCollection = New clsCollection
		    		  	
		    		  objCollection.CollectionType = objCollection.clContact
		    		  objCollection.Query = "SELECT DISTINCT tblContacts.Username, tblContacts.ContactPK FROM tblContacts " &_
		    		  					            "INNER JOIN tblGroupMembers ON tblContacts.ContactPK = tblGroupMembers.ContactFK " & _
		    		  					            "WHERE tblContacts.IsActive=" & lhd_True & " " &_
		    		  					            "ORDER BY tblContacts.UserName ASC"
		    		  							 
		    		  If Not objCollection.Load Then
		    		  		
		    		      Response.Write objCollection.LastError
		    		  		    
		    		  Else
		    		  			
		    		  	If objCollection.BOF And objCollection.EOF Then
		    		  			
		    		  		' No records
		    		  			
		    		  	Else
		    		  		
		    		  		Do While Not objCollection.EOF
		    		  		%>
		    		  			<OPTION VALUE="<%=objCollection.Item.ID%>"><%=objCollection.Item.UserName%></OPTION>
		    		  		<%
		    		  			objCollection.MoveNext
		    		  		Loop
		    		  				
		    		  	End If
		    		  			
		    		  End If
		    		  	
		    		  Set objCollection = Nothing
		    		  %>
		    		</SELECT>
		    	</TD>
		      <TD></TD>
		    	<TD></TD>
		    	<TD></TD>
		    	<TD></TD>
		    </TR>
		    <TR>
		    	<TD colspan="7"></TD>
		    </TR>
		    <TR class="lhd_Heading2">
		    	<TD colspan="7"><%=Lang("Date_Criteria")%></TD>
		    </TR>
        <TR>
          <TD></TD>
          <TD><INPUT type="radio" id="radAllAvailable" name=radio2 checked>&nbsp;<%=Lang("All")%></TD>
          <TD></TD>
          <TD></TD>
          <TD></TD>
          <TD></TD>
          <TD></TD>
        </TR>
        <TR>
          <TD></TD>
          <TD><INPUT type="radio" id="radCustom" name=radio2>&nbsp;<%=Lang("Custom")%></TD>
          <TD align="Right">... <%=Lang("From")%>:&nbsp;&nbsp;<INPUT style="WIDTH=95px;" type="text" id=tbxDateFrom name=tbxDateFrom>&nbsp;<A href="javascript: ;" OnClick="javascript:makeCalendar('document.frmSearch.tbxDateFrom', 'document.frmSearch.tbxDateFrom.value') ;"><IMG src="Images/Calendar.gif" border=0 ALIGN="absmiddle"></A></TD>
          <TD></TD>
          <TD></TD>
          <TD></TD>
          <TD></TD>
        </TR>
        <TR>
          <TD></TD>
          <TD></TD>
          <TD align="Right">... <%=Lang("To")%>:&nbsp;&nbsp;<INPUT style="WIDTH=95px;" type="text" id=tbxDateTo name=tbxDateTo>&nbsp;<A href="javascript: ;" OnClick="javascript:makeCalendar('document.frmSearch.tbxDateTo', 'document.frmSearch.tbxDateTo.value') ;"><IMG src="Images/Calendar.gif" border=0 ALIGN="absmiddle"></A> </TD>
          <TD></TD>
          <TD></TD>
          <TD></TD>
          <TD></TD>
        </TR>
		    <TR>
		      <TD colspan="7"></TD>
		    </TR>
		    <TR>
		      <TD align="Right" colspan="6"><INPUT style="WIDTH: 80px; BACKGROUND-COLOR: white" id=btnSearch name=btnSearch type=submit value="<%=Lang("Search")%>"></TD>
          <TD></TD>
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
