<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: clsDepartment.asp
'  Date:     $Date: 2004/03/11 06:29:08 $
'  Version:  $Revision: 1.2 $
'  Purpose:  Class object to assist when working with Departments
' ----------------------------------------------------------------------------------

Class clsDepartment
  ' ##################
  ' Private Properties
  ' ##################
  Private m_ID, m_Org, m_OrgID, m_IsActive
  Private m_DeptName, m_DeptDesc
  Private m_LastUpdate, m_LastUpdateBy, m_LastUpdateByID

  Private m_IsValid, m_LastError, m_cnnDB

  ' #################
  ' Public Properties
  ' #################
  Public SourceRS ' As RecordSet

  Public Property Get ID()  ' As Long
    ID = m_ID
  End Property

  Public Property Let ID(f_ID)
    If IsNumeric(f_ID) Then
      m_ID = f_ID
    End If
  End Property

  Public Property Get Org() ' As clsOrganisation
    If Not IsObject(m_Org) Then
      Dim oOrg, blnTemp
      Set oOrg = New clsOrganisation
      If IsNumeric(m_OrgID) Then
        oOrg.ID = m_OrgID
        blnTemp = oOrg.Load  ' Ignore Errors
      End If
      Set m_Org = oOrg
    End If
    Set Org = m_Org
  End Property

  Public Property Get OrgID() ' As Long
    OrgID = m_OrgID    
  End Property

  Public Property Let OrgID(ByVal f_OrgID)
    Dim oOrg
    Set oOrg = New clsOrganisation
    oOrg.ID = f_OrgID
    If Not oOrg.Load Then
      ' ########
      ' Raise Error
    End If
    m_OrgID = f_OrgID
    Set m_Org = oOrg
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

  Public Property Get DeptName()  ' As String
    DeptName = m_DeptName
  End Property

  Public Property Let DeptName(ByRef f_DeptName)
    m_DeptName = Left(Trim(f_DeptName), 36)
  End Property

  Public Property Get DeptDesc()  ' As String
    DeptDesc = m_DeptDesc
  End Property

  Public Property Let DeptDesc(ByRef f_DeptDesc)
    m_DeptDesc = Left(Trim(f_DeptDesc), 160)
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

    Dim blnLoad, rsDept, strQuery, blnUseSourceRS
    blnLoad = True
    blnUseSourceRS = False
    If IsObject(SourceRS) Then
      If SourceRS.State <> adStateClosed Then
        Set rsDept = SourceRS
        blnUseSourceRS = True
      Else
        blnLoad = False
        m_LastError = "SourceRS is closed."
      End If
    Else
      If IsNumeric(m_ID) And Not IsEmpty(m_ID) Then
        strQuery = "SELECT * FROM tblDepartments WHERE DeptPK = " & m_ID
        Set rsDept = Server.CreateObject("ADODB.RecordSet")
        rsDept.Open strQuery, m_cnnDB
      Else
        blnLoad = False
        m_LastError = "Missing ID."
      End If
    End If
    If blnLoad Then
      If Not rsDept.EOF And Not rsDept.BOF Then
        On Error Resume Next
        m_ID = rsDept("DeptPK")
        m_OrgID = rsDept("OrgFK")
        m_IsActive = rsDept("IsActive")
        m_DeptName = rsDept("DeptName")
        m_DeptDesc = rsDept("DeptDesc")
        m_LastUpdate = rsDept("LastUpdate")
        m_LastUpdateBy = rsDept("LastUpdateByFK")
        On Error Goto 0
        If Err.Number <> 0 Then
          blnLoad = False
          m_LastError = "Error retrieving data -- " & Err.Source & ": " & Err.Description
        End If
      Else
        blnLoad = False
        m_LastError = "No matching departments found."
      End If
    End If
    If Not blnUseSourceRS Then
      rsDept.Close
      Set rsDept = Nothing
    End If
    If blnLoad Then
      m_IsValid = True
    End If
    Load = blnLoad
  End Function

  Public Function Update() ' As Boolean
    Dim blnUpdate, rsDept
    blnUpdate = True
    Set rsDept = Server.CreateObject("ADODB.RecordSet")
    If m_IsValid Then
      rsDept.Open "SELECT * FROM tblDepartments WHERE DeptPK = " & m_ID, m_cnnDB, _
        adOpenKeyset, adLockOptimistic, adCmdText
    Else
      rsDept.Open "tblDepartments", m_cnnDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect
      rsDept.AddNew
    End If
    With rsDept
      .Fields("OrgFK") = m_OrgID
      .Fields("IsActive") = m_IsActive
      .Fields("DeptName") = m_DeptName
      .Fields("DeptDesc") = m_DeptDesc
      .Fields("LastUpdate") = m_LastUpdate
      .Fields("LastUpdateByFK") = m_LastUpdateByID
      .Update
      m_ID = .Fields("DeptPK")
    End With
    rsDept.Close
    Set rsDept = Nothing
    m_IsValid = True
    Update = blnUpdate
  End Function

  Public Function Delete() ' As Boolean
    Dim blnDelete
    blnDelete = True
    m_IsActive = False
    If m_IsValid Then
      m_cnnDB.Execute "UPDATE tblDepartments SET IsActive = " & m_IsActive & _
        " WHERE DeptPK = " & m_ID, , adExecuteNoRecords
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

    ' Set Default Contact Properties
'    m_IsActive = True
'     m_DeptName = "Unknown"
  End Sub

  Private Sub Class_Terminate()
  End Sub

End Class
%>