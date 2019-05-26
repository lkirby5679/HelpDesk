<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: clsGroup.asp
'  Date:     $Date: 2004/03/11 06:29:08 $
'  Version:  $Revision: 1.2 $
'  Purpose:  Class object to assist when working with Groups
' ----------------------------------------------------------------------------------

Class clsGroup
  ' #################
  ' Private Poperties
  ' #################
  Private m_ID, m_GroupName, m_GroupDesc
  Private m_GroupMembers, m_IsActive, m_LastUpdate, m_LastUpdateBy, m_LastUpdateByID

  Private m_IsValid, m_LastError, m_cnnDB
  ' #################
  ' Public Poperties
  ' #################
  Public SourceRS ' As RecordSet

  Public Property Get ID()  ' As Long
    ID = m_ID
  End Property

  Public Property Let ID(ByRef f_ID)
    If IsNumeric(f_ID) Then
      m_ID = f_ID
    Else
      ' Invalid ID
    End If
  End Property

  Public Property Get GroupName()  ' As String
    GroupName = m_GroupName
  End Property

  Public Property Let GroupName(ByRef f_GroupName)
    m_GroupName = Left(Trim(f_GroupName), 32)
  End Property

  Public Property Get GroupDesc()  ' As String
    GroupDesc = m_GroupDesc
  End Property

  Public Property Let GroupDesc(ByRef f_GroupDesc)
    m_GroupDesc = Left(Trim(f_GroupDesc), 160)
  End Property

  Public Property Get GroupMembers()  ' As clsCollection of Contacts
  
    Dim objGroupMembers
  
    If m_IsValid Then

      If Not IsObject(m_GroupMembers) Then
      
        Set objGroupMembers = New clsCollection

        objGroupMembers.CollectionType = objGroupMembers.clContact
        objGroupMembers.Query = "SELECT tblContacts.* FROM tblContacts " & _
          "INNER JOIN tblGroupMembers ON (tblContacts.ContactPK = tblGroupMembers.ContactFK) " & _
          "WHERE tblGroupMembers.GroupFK = " & m_ID & " ORDER BY tblContacts.UserName ASC"
          
        If Not objGroupMembers.Load Then
		  ' Raise Error
		Else
		  Set m_GroupMembers = objGroupMembers
		End If
		
		Set objGroupMembers = Nothing
		
	  Else
	  
	    ' m_GroupMembers already loaded

      End If

      Set GroupMembers = m_GroupMembers

    Else
    
      ' Raise error as group invalid
      
    End If
    
  End Property

  Public Property Get IsActive() ' As Boolean
    IsActive = m_IsActive
  End Property

  Public Property Let IsActive(f_IsActive) ' As Boolean
    If IsNumeric(f_IsActive) Then
      m_IsActive = f_IsActive
    End If
    IsActive = m_IsActive
  End Property

  Public Property Get LastUpdate() ' As Date
    LastUpdate = m_LastUpdate
  End Property

  Public Property Let LastUpdate(ByRef f_dtLastUpdate)
    If Not IsDate(f_dtLastUpdate) Then
      ' #######
      ' Print an error and die
    End If
    m_LastUpdate = f_dtLastUpdate
  End Property

  Public Property Get LastUpdateBy() ' As clsContact
    If Not IsObject(m_LastUpdateBy) Then
      Dim oContact, blnTemp
      Set oContact = New clsContact
      If IsNumeric(m_LastUpdateByID) Then
        oContact.ID = m_LastUpdateByID
        blnTemp = oContact.Load
      End If
      Set m_LastUpdateBy = oContact
    End If
  Set LastUpdateBy = m_LastUpdateBy
  End Property

  Public Property Get LastUpdateByID() ' As Long
    LastUpdateByID = m_LastUpdateByID
  End Property

  Public Property Let LastUpdateByID(f_LastUpdateByID) ' As Long
    m_LastUpdateByID = f_LastUpdateByID
  End Property

  Public Property Get IsValid() ' As Boolean
    IsValid = m_IsValid
  End Property

  Public Property Get LastError() ' As String
    LastError = m_LastError
  End Property


  ' ##############
  ' Public Methods
  ' ##############
  Public Function Load()  ' As Boolean
    ' One of the following is required to load:
    ' SourceRS
    ' ID

    Dim blnLoad, rsGroup, strQuery, blnUseSourceRS
    blnLoad = True
    blnUseSourceRS = False
    If IsObject(SourceRS) Then
      If SourceRS.State <> adStateClosed Then
        Set rsGroup = SourceRS
        blnUseSourceRS = True
      Else
        blnLoad = False
        m_LastError = "SourceRS is closed."
      End If
    Else
      If IsNumeric(m_ID) And Not IsEmpty(m_ID) Then
        strQuery = "SELECT * FROM tblGroups WHERE GroupPK = " & m_ID
        Set rsGroup = Server.CreateObject("ADODB.RecordSet")
        rsGroup.Open strQuery, m_cnnDB
      Else
        blnLoad = False
        m_LastError = "Missing ID."
      End If
    End If
    If blnLoad Then
      If Not rsGroup.EOF And Not rsGroup.BOF Then
        On Error Resume Next
        m_ID = rsGroup("GroupPK")
        m_GroupName = rsGroup("GroupName")
        m_GroupDesc = rsGroup("GroupDesc")
        m_RoleMask = rsGroup("RoleMask")
        m_IsActive = rsGroup("IsActive")
        m_LastUpdate = rsGroup("LastUpdate")
        m_LastUpdateBy = rsGroup("LastUpdateByFK")
        On Error Goto 0
        If Err.Number <> 0 Then
          blnLoad = False
          m_LastError = "Error retrieving data -- " & Err.Source & ": " & Err.Description
        End If
      Else
        blnLoad = False
        m_LastError = "No matching groups found."
      End If
    End If
    If Not blnUseSourceRS And blnLoad Then
      rsGroup.Close
      Set rsGroup = Nothing
    End If
    If blnLoad Then
      m_IsValid = True
    End If
    Load = blnLoad
  End Function

  Public Function Update() ' As Boolean
    Dim blnUpdate, rsGroup
    blnUpdate = True
    Set rsGroup = Server.CreateObject("ADODB.RecordSet")
    If m_IsValid Then
      rsGroup.Open "SELECT * FROM tblGroups WHERE GroupPK = " & m_ID, m_cnnDB, _
        adOpenKeyset, adLockOptimistic, adCmdText
    Else
      rsGroup.Open "tblGroups", m_cnnDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect
      rsGroup.AddNew
    End If
    With rsGroup
      .Fields("GroupName") = m_GroupName
      .Fields("GroupDesc") = m_GroupDesc
      .Fields("IsActive") = m_IsActive
      .Fields("LastUpdate") = m_LastUpdate
      .Fields("LastUpdateByFK") = m_LastUpdateByID
      .Update
      m_ID = .Fields("GroupPK")
    End With
    rsGroup.Close
    Set rsGroup = Nothing
    m_IsValid = True
    Update = blnUpdate
  End Function

  Public Function Delete() ' As Boolean
    Dim blnDelete
    blnDelete = True
    m_IsActive = False
    If m_IsValid Then
      m_cnnDB.Execute "UPDATE tblGroups SET IsActive = " & m_IsActive & _
        " WHERE GroupPK = " & m_ID, , adExecuteNoRecords
    End If
    Delete = blnDelete
  End Function

	Public Function IsMember(ByVal f_MemberID)	' As Boolean
	
		Dim rstMember
		
	
		Set rstMember = Server.CreateObject("ADODB.Recordset")
			
		rstMember.Open "SELECT * FROM tblGroupMembers WHERE GroupFK=" & m_ID & " AND ContactFK=" & f_MemberID, cnnDB
		
		If rstMember.BOF And rstMember.EOF Then
			IsMember = False
		Else
			IsMember = True
		End If
			
		rstMember.Close
			
		Set rstMember = Nothing
	
	End Function
	


  Public Function AddMember(f_MemberID) ' As Boolean

    Dim blnMemberAdded
    Dim blnIsMember
    Dim rsMember
    
    
    m_IsActive = False
    blnIsMember = False
    blnMemberAdded = False
    
    If m_IsValid Then
		
	  '  First need to determine if f_MemberID is already a member. If so then we
	  ' remove him
	  
	  Set rsMember = Server.CreateObject("ADODB.Recordset")
	  rsMember.Open "SELECT ContactFK FROM tblGroupMembers WHERE GroupFK=" & m_ID & " AND ContactFK=" & f_MemberID, m_cnnDB

	  If rsMember.BOF And rsMember.EOF Then
		blnIsMember = False
	  Else
		blnIsMember = True
	  End If
	  
	  rsMember.Close
	  Set rsMember = Nothing
	  
	  
	  If blnIsMember = False Then    
	  
        m_cnnDB.Execute "INSERT INTO tblGroupMembers(GroupFK, ContactFK, LastUpdate, LastUpdateByFK) " &_
						"VALUES(" & m_ID & ", " & f_MemberID & ", '" & m_LastUpdate & "', " & m_LastUpdateByID & ")"
        
        blnMemberAdded = True

      Else
      
        blnMemberAdded = False
        
      End If
      
    End If
    
    AddMember = blnMemberAdded

  End Function


  Public Function RemoveMember(f_MemberID) ' As Boolean

    Dim blnMemberRemoved
    Dim blnIsMember
    Dim rsMember
    
    
    m_IsActive = False
    blnIsMember = False
    blnMemberRemoved = False
    
    If m_IsValid Then
		
	  '  First need to determine if f_MemberID is already a member. If so then we
	  ' remove him
	  
	  Set rsMember = Server.CreateObject("ADODB.Recordset")
	  rsMember.Open "SELECT ContactFK FROM tblGroupMembers WHERE GroupFK=" & m_ID & " AND ContactFK=" & f_MemberID, m_cnnDB

	  If rsMember.BOF And rsMember.EOF Then
		blnIsMember = False
	  Else
		blnIsMember = True
	  End If
	  
	  rsMember.Close
	  Set rsMember = Nothing
	  
	  
	  If blnIsMember = True Then    
	  
        m_cnnDB.Execute "DELETE FROM tblGroupMembers " &_
						"WHERE GroupFK=" & m_ID & " AND ContactFK=" & f_MemberID
        
        blnMemberRemoved = True

      Else
      
        blnMemberRemoved = False
        
      End If
      
    End If
    
    RemoveMember = blnMemberRemoved

  End Function


  ' ###############
  ' Private Methods
  ' ###############
  Private Sub Class_Initialize()
    If IsObject(cnnDB) Then
      Set m_cnnDB = cnnDB
    End If
    m_IsValid = False

    ' Set Default Contact Properties
'    m_IsActive = True
    m_GroupName = "-"
  End Sub

  Private Sub Class_Terminate()
  End Sub
End Class
%>