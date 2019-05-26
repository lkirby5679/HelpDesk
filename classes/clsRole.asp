<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: clsRole.asp
'  Date:     $Date: 2004/03/11 06:29:08 $
'  Version:  $Revision: 1.2 $
'  Purpose:  Class object to assist when working with Roles
' ----------------------------------------------------------------------------------

Class clsRole
  ' ##################
  ' Private Properties
  ' ##################
  Private m_ID, m_RoleName, m_RoleDesc, m_RoleMask, m_IsActive
  Private m_LastUpdate, m_LastUpdateBy, m_LastUpdateByID

  Private m_IsValid, m_LastError, m_cnnDB

  ' #################
  ' Public Properties
  ' #################
  Public SourceRS ' As RecordSet

  Public Property Get ID()  ' As Long
    ID = m_ID
  End Property

  Public Property Let ID(ByRef f_ID)
    If IsNumeric(f_ID) Then
      m_ID = f_ID
    End If
  End Property

  Public Property Get RoleName()  ' As String
    RoleName = m_RoleName
  End Property

  Public Property Let RoleName(ByRef f_RoleName)
    m_RoleName = Left(Trim(f_RoleName), 80)
  End Property

  Public Property Get RoleDesc()  ' As String
    RoleDesc = m_RoleDesc
  End Property

  Public Property Let RoleDesc(ByRef f_RoleDesc)
    m_RoleDesc = Left(Trim(f_RoleDesc), 255)
  End Property

  Public Property Let RoleMask(ByRef f_RoleMask)
    m_RoleMask = CLng(f_RoleMask) OR &H0000000000000000
    'Response.Write m_RoleMask
'    m_RoleMask = f_RoleMask
  End Property

  Public Property Get RoleMask()  ' As Binary
    RoleMask = m_RoleMask
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

  Public Property Get IsActive() ' As Boolean
    IsActive = m_IsActive
  End Property

  Public Property Let IsActive(f_IsActive) ' As Boolean
    If IsNumeric(f_IsActive) Then
		m_IsActive=f_IsActive
    End If
    IsActive = m_IsActive
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

    Dim blnLoad, rsRole, strQuery, blnUseSourceRS
    blnLoad = True
    blnUseSourceRS = False
    If IsObject(SourceRS) Then
      If SourceRS.State <> adStateClosed Then
        Set rsRole = SourceRS
        blnUseSourceRS = True
      Else
        blnLoad = False
        m_LastError = "SourceRS is closed."
      End If
    Else
      If IsNumeric(m_ID) Then
'        strQuery = "SELECT RoleName, RoleDesc, CONVERT(INT, RoleMask) AS 'RoleMask', IsActive, LastUpdate, LastUpdateByFK FROM tblRoles WHERE RolePK = " & m_ID
        strQuery = "SELECT RoleName, RoleDesc, RoleMask, IsActive, LastUpdate, LastUpdateByFK FROM tblRoles WHERE RolePK = " & m_ID
        Set rsRole = Server.CreateObject("ADODB.RecordSet")
        rsRole.Open strQuery, m_cnnDB
      Else
        blnLoad = False
        m_LastError = "Missing ID."
      End If
    End If
    If blnLoad Then
      If Not rsRole.EOF And Not rsRole.BOF Then
        On Error Resume Next
        m_ID = rsRole("RolePK")
        m_RoleName = rsRole("RoleName")
        m_RoleDesc = rsRole("RoleDesc")
        m_RoleMask = rsRole("RoleMask")
        m_IsActive = rsRole("IsActive")
        m_LastUpdate = rsRole("LastUpdate")
        m_LastUpdateBy = rsRole("LastUpdateByFK")
        On Error Goto 0
        If Err.Number <> 0 Then
          blnLoad = False
          m_LastError = "Error retrieving data -- " & Err.Source & ": " & Err.Description
        End If
      Else
        blnLoad = False
        m_LastError = "No matching roles found."
      End If
    End If
    If Not blnUseSourceRS Then
      rsRole.Close
      Set rsRole = Nothing
    End If
    If blnLoad Then
      m_IsValid = True
    End If
    Load = blnLoad
  End Function

  Public Function Update() ' As Boolean
    Dim blnUpdate, rsRole, strSQL
    blnUpdate = True
    Set rsRole = Server.CreateObject("ADODB.RecordSet")
    If m_IsValid Then
      rsRole.Open "SELECT * FROM tblRoles WHERE RolePK = " & m_ID, m_cnnDB, _
        adOpenKeyset, adLockOptimistic, adCmdText

'     strSQL = "UPDATE tblRoles SET RoleName='" & m_RoleName & "', RoleDesc='" & m_RoleDesc & "', RoleMask=" & m_RoleMask & ", "
'	 strSQL = strSQL & "IsActive=" & m_IsActive & ",LastUpdate='" & m_LastUpdate & "', LastUpdateByFK=" & m_LastUpdateByID & ") "
'     strSQL = strSQL & "WHERE RolePK=" & m_ID

'     rsRole.Open strSQL, m_cnnDB

    Else
      rsRole.Open "tblRoles", m_cnnDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect
      rsRole.AddNew
      
'     strSQL = "INSERT INTO tblRoles (RoleName, RoleDesc, RoleMask, IsActive, LastUpdate, LastUpdateByFK) "
'     strSQL = strSQL & "VALUES ('" & m_RoleName & "','" & m_RoleDesc & "'," & m_RoleMask & "," & m_IsActive & ",'" & m_LastUpdate & "'," &  m_LastUpdateByID & ")"
   
'     rsRole.Open strSQL, m_cnnDB
      
    End If

	With rsRole
	  .Fields("RoleName") = m_RoleName
	  .Fields("RoleDesc") = m_RoleDesc
	  .Fields("RoleMask") = m_RoleMask
	  .Fields("IsActive") = m_IsActive
	  .Fields("LastUpdate") = m_LastUpdate
	  .Fields("LastUpdateByFK") = m_LastUpdateByID
	  .Update
	End With
    m_ID = rsRole.Fields("RolePK")

    rsRole.Close

'	' Update the RoleMask separately
'
'    strSQL = "UPDATE tblRoles SET RoleMask=" & m_RoleMask & " WHERE RolePK=" & m_ID
'    rsRole.Open strSQL, m_cnnDB
    
    Set rsRole = Nothing
    m_IsValid = True
    Update = blnUpdate
  End Function

  Public Function Delete() ' As Boolean
    Dim blnDelete
    blnDelete = True
    m_IsActive = False
    If m_IsValid Then
      m_cnnDB.Execute "UPDATE tblRoles SET IsActive = " & m_IsActive & _
        " WHERE RolePK = " & m_ID, , adExecuteNoRecords
    End If
    Delete = blnDelete
  End Function

  ' ###############
  ' Private Methods
  ' ###############
  Private Sub Class_Initialize()
    If IsObject(cnnDB) Then
      Set m_cnnDB = cnnDB
    End If
    m_IsValid = False

'    ' Set Default Contact Properties
 '   m_IsActive = True
  '  m_RoleName = "Unknown"
   ' m_RoleMask = &H0000000000000000
  End Sub

  Private Sub Class_Terminate()
  End Sub
End Class
%>