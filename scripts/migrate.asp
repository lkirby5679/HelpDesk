<% 

Option Explicit

Response.Buffer = False

Server.ScriptTimeout = 600

%>
<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: Migrate.asp
'  Date:     $Date: 2004/03/30 06:11:39 $
'  Version:  $Revision: 1.15 $
'  Purpose:  Used in assistin gthe migration of data from Liberum 97.2/3 to Huron
' ----------------------------------------------------------------------------------
%>

<html>

<!--METADATA TYPE="TypeLib" NAME="Microsoft ActiveX Data Objects 2.5 Library" UUID="{00000205-0000-0010-8000-00AA006D2EA4}" VERSION="2.5"-->
<!--METADATA TYPE="TypeLib" NAME="Microsoft Scripting Runtime" UUID="{420B2830-E718-11CF-893D-00A0C9054228}" VERSION="1.0"-->
<!--METADATA TYPE="TypeLib" NAME="Microsoft CDO for Windows 2000 Library" UUID="{CD000000-8B95-11D1-82DB-00C04FB1625D}" VERSION="1.0"-->

<!-- #Include File = "Include/Settings.asp" -->
<!-- #Include File = "Include/Public.asp" -->

<!-- #Include File = "Classes/clsAssignment.asp" -->
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

<body>


<%

Dim cnnSource, cnnDB
Dim rstSource, rstDestination
Dim strSQL
Dim objContact, objCase, objCategory, objDepartment
Dim blnMigrateUsers, blnMigrateCases, blnMigrateCategories, blnMigrateDepartments



blnMigrateDepartments = True
blnMigrateUsers = True
blnMigrateCategories = True
blnMigrateCases = True

' Create connection to the database

Set cnnSource = Server.CreateObject("ADODB.Connection")
cnnSource.Open "Provider=SQLOLEDB.1;Data Source=AA000891;Initial Catalog=SUPPORT;uid=sa;pwd=password"
'cnnSource.Open "Provider=SQLOLEDB.1;Data Source=COUNT;Initial Catalog=SUPPORT";uid=" & Application("SQLUser") & ";pwd=" & Application("SQLPass")

Set cnnDB = Server.CreateObject("ADODB.Connection")
cnnDB.Open "Provider=SQLOLEDB.1;Data Source=AA000891;Initial Catalog=HURON;uid=sa;pwd=password"
'cnnDB.Open "Provider=SQLOLEDB.1;Data Source=COUNT;Initial Catalog=HURON;uid=" & Application("SQLUser") & ";pwd=" & Application("SQLPass")

'
' ----------------------------------------------------------------------------------------------
'
' Migrate "departments" Table

Response.Write "Processing Departments "

If blnMigrateDepartments = True Then 

  Set rstSource = Server.CreateObject("ADODB.Recordset")
  rstSource.Open "SELECT * FROM departments ORDER BY department_id ASC", cnnSource
  						    
  If rstSource.BOF And rstSource.EOF Then
  	  					    
  	' No records
  	  						
  Else
  	
  	While Not rstSource.EOF
  	
  	  Set objDepartment = New clsDepartment
  	  
  	  objDepartment.ID = rstSource.Fields("department_id").Value 
  	  
  	  If Not objDepartment.Load Then
  	  
  	    ' Contact doesn't not exist so we must create it.
  	    
  	    objDepartment.DeptName = rstSource.Fields("dname").Value
  	    objDepartment.IsActive = 1
  	    
  	    If Not objDepartment.Update Then
  	    
  	      ' Failed to created record
  	      
  	    Else
  	    
  	      ' Record created
  	      
  	    End If
  	  
  	  Else
  	  
  	    ' Do nothing, as contact exists.
  	  
  	  End If
  	  
      Set objDepartment = Nothing
  	  			
      Response.Write " ."
  	  	
  		rstSource.MoveNext
      
  	WEnd

  End If

  rstSource.Close 
  Set rstSource = Nothing

Else

  ' Do nothing

End If

Response.Write " Complete<BR><BR>"

'
' ----------------------------------------------------------------------------------------------
'
' Migrate "tblUsers" Table

Response.Write "Processing Contacts/Users "
  	  	
If blnMigrateUsers = True Then 

  Set rstSource = Server.CreateObject("ADODB.Recordset")
  rstSource.Open "SELECT * FROM tblUsers ORDER BY sid ASC", cnnSource
  						    
  If rstSource.BOF And rstSource.EOF Then
  	  					    
  	' No records
  	  						
  Else
  	
  	While Not rstSource.EOF
  	
  	  Set objContact = New clsContact
  	  
  	  objContact.UserName = rstSource.Fields("uid").Value 
  	  
  	  If Not objContact.Load Then
  	  
  	    ' Contact doesn't not exist so we must create it.
  	    
  	    objContact.UserName = rstSource.Fields("uid").Value
  	    objContact.FName = rstSource.Fields("firstname").Value
  	    objContact.LName = rstSource.Fields("lastname").Value
  	    objContact.Email = rstSource.Fields("email1").Value
  	    objContact.OfficeLocation = rstSource.Fields("location1").Value
  	    objContact.HomePhone = rstSource.Fields("phone_home").Value
  	    objContact.OfficePhone = rstSource.Fields("phone").Value
  	    objContact.MobilePhone = rstSource.Fields("phone_mobile").Value
  	    objContact.RoleID = 3
  	    objContact.LangID = 1
        objContact.OrgID = 1
        objContact.DeptID = rstSource.Fields("department").Value
  	    
  	    
  	    If Not objContact.Update Then
  	    
  	      ' Failed to created record
  	      
  	    Else
  	    
  	      ' Record created
  	      
  	    End If
  	  
  	  Else
  	  
  	    ' Do nothing, as contact exists.
  	  
  	  End If
  	  
      Set objContact = Nothing
  	  			
      Response.Write " ."
  	  	
  		rstSource.MoveNext
  	  	
  	WEnd

  End If

  rstSource.Close 
  Set rstSource = Nothing

Else

  ' Do nothing

End If

Response.Write " Complete<BR><BR>"

'
' ----------------------------------------------------------------------------------------------
'
' Migrate "categories" Table

Response.Write "Processing Categories "
  	  	
If blnMigrateCategories = True Then 

  Set rstSource = Server.CreateObject("ADODB.Recordset")
  rstSource.Open "SELECT * FROM categories ORDER BY category_id ASC", cnnSource
  						    
  If rstSource.BOF And rstSource.EOF Then
  	  					    
  	' No records
  	  						
  Else
  	
  	While Not rstSource.EOF
  	
  	  Set objCategory = New clsCategory
  	  
  	  objCategory.ID = rstSource.Fields("category_id").Value 
  	  
  	  If Not objCategory.Load Then
  	  
  	    ' Contact doesn't not exist so we must create it.
  	    
  	    objCategory.CaseTypeID = 1
  	    objCategory.CatName = rstSource.Fields("cname").Value
  	    objCategory.IsActive = 1
  	    
  	    If Not objCategory.Update Then
  	    
  	      ' Failed to created record
  	      
  	    Else
  	    
  	      ' Record created
  	      
  	    End If
  	  
  	  Else
  	  
  	    ' Do nothing, as contact exists.
  	  
  	  End If
  	  
      Set objCategory = Nothing
  	  			
      Response.Write " ."
  	  	
  		rstSource.MoveNext
  	  	
  	WEnd

  End If

  rstSource.Close 
  Set rstSource = Nothing

Else

  ' Do nothing

End If

Response.Write " Complete<BR><BR>"

'
' ----------------------------------------------------------------------------------------------
'
' Migrate "problems" Table

Response.Write "Processing Cases "
  	  	
If blnMigrateCases = True Then 

  Set rstSource = Server.CreateObject("ADODB.Recordset")
  rstSource.Open "SELECT problems.*, tblUsers.uid AS RepName FROM problems INNER JOIN tblUsers ON problems.rep = tblUsers.sid ORDER BY id ASC", cnnSource
  						    
  If rstSource.BOF And rstSource.EOF Then
  	  					    
  	' No records
  	  						
  Else
  	
  	While Not rstSource.EOF
  	
  	  Set objCase = New clsCase
  	  
  	  objCase.ID = rstSource.Fields("id").Value 
  	  
  	  If Not objCase.Load Then
  	  
  	    ' Contact doesn't not exist so we must create it.
  	    
  	    Set objContact = New clsContact
  	    
  	    objContact.UserName = rstSource.Fields("uid").Value
  	    
        If Not objContact.Load Then
          ' No record found
        Else
    	    objCase.ContactID = objContact.ID
        End If  	   
  	    
  	    Set objContact = Nothing
  	    
  	    objCase.DeptID = rstSource.Fields("department").Value
  	    objCase.CaseTypeID = 1
  	    objCase.CatID = rstSource.Fields("category").Value

  	    Select Case rstSource.Fields("priority").Value
  	    
  	      Case 1  ' Low
  	        objCase.PriorityID = 7

  	      Case 2  ' Normal
  	        objCase.PriorityID = 8

  	      Case 3  ' High
  	        objCase.PriorityID = 9

  	      Case 4  ' Urgent
  	        objCase.PriorityID = 10
  	    
  	      Case Else
  	        objCase.PriorityID = 8
  	        
  	    End Select
  	    
  	    Select Case rstSource.Fields("status").Value
  	      
  	      Case 1  ' Open
     	      objCase.StatusID = 2
     	      
  	      Case 10  ' Actioning
     	      objCase.StatusID = 3

  	      Case 15  ' On Hold
     	      objCase.StatusID = 3

  	      Case 20  ' Awaiting Repairs
     	      objCase.StatusID = 3

  	      Case 25  ' Awaiting Response
     	      objCase.StatusID = 3
  	      
  	      Case 100  ' Closed
     	      objCase.StatusID = 5
  	    
          Case Else	    
     	      objCase.StatusID = 3
     	      
  	    End Select
  	    

  	    Set objContact = New clsContact
  	    
  	    objContact.UserName = rstSource.Fields("RepName").Value
  	    
        If Not objContact.Load Then
          ' No record found
        Else
    	    objCase.RepID = objContact.ID
        End If  	   
  	    
  	    Set objContact = Nothing

  	    objCase.RaisedDate = rstSource.Fields("start_date").Value
  	    objCase.ClosedDate = rstSource.Fields("close_date").Value
  	    objCase.GroupID = 1
  	    objCase.EnteredByID = rstSource.Fields("entered_by").Value
  	    objCase.LastUpdate = rstSource.Fields("modified_date").Value
  	    objCase.Title = rstSource.Fields("title").Value
  	    objCase.Description = rstSource.Fields("description").Value
  	    objCase.Resolution = rstSource.Fields("solution").Value
  	    objCase.IsActive = 1
  	    
  	    If Not objCase.Update Then
  	    
  	      ' Failed to created record
  	      
  	    Else
  	    
  	      ' Record created
  	      
  	    End If
  	  
  	  Else
  	  
  	    ' Do nothing, as contact exists.
  	  
  	  End If
  	  
      Set objCase = Nothing
  	  			
      Response.Write " ."
  	  	
  		rstSource.MoveNext
  	  	
  	WEnd

  End If

  rstSource.Close 
  Set rstSource = Nothing

Else

  ' Do nothing

End If

cnnSource.Close
cnnDB.Close 

Set cnnSource = Nothing
Set cnnDB = Nothing

Response.Write " Complete<BR><BR>"

%>


</body>


</html>