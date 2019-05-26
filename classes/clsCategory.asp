<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: clsCategory.asp
'  Date:     $Date: 2004/03/11 06:29:08 $
'  Version:  $Revision: 1.2 $
'  Purpose:  Class object to assist when working with Categories
' ----------------------------------------------------------------------------------

Class clsCategory
  ' ##################
  ' Private Properties
  ' ##################
  Private m_ID, m_Org, m_IsActive
  Private m_CatName, m_CatDesc, m_CatOrder, m_CaseType, m_CaseTypeID
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

  Public Property Get CaseType() ' As clsCaseType
    If Not IsObject(m_CaseType) Then
      Dim oCaseType, blnTemp
      Set oCaseType = New clsCaseType
      If IsNumeric(m_CaseTypeID) Then
        oCaseType.ID = m_CaseTypeID
        blnTemp = oCaseType.Load
      End If
      Set m_CaseType = oCaseType
    End If
    Set CaseType = m_CaseType
  End Property

  Public Property Get CaseTypeID() ' As Long
    CaseTypeID = m_CaseTypeID
  End Property

  Public Property Let CaseTypeID(ByRef f_CaseTypeID)
    Dim oCaseType
    Set oCaseType = New clsCaseType
    oCaseType.ID = f_CaseTypeID
    If Not oCaseType.Load Then
      ' ######
      ' Raise Error
    Else
      m_CaseTypeID = f_CaseTypeID
      Set m_CaseType = oCaseType
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

  Public Property Get CatName()  ' As String
    CatName = m_CatName
  End Property

  Public Property Let CatName(ByRef f_CatName)
    m_CatName = Left(Trim(f_CatName), 36)
  End Property

  Public Property Get CatDesc()  ' As String
    CatDesc = m_CatDesc
  End Property

  Public Property Let CatDesc(ByRef f_CatDesc)
    m_CatDesc = Left(Trim(f_CatDesc), 160)
  End Property

  Public Property Get CatOrder()  ' As String
    CatOrder = m_CatOrder
  End Property

  Public Property Let CatOrder(ByRef f_CatOrder)
    m_CatOrder = f_CatOrder
  End Property

  Public Property Get IsValid() ' As Boolean
    IsValid = m_IsValid
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

    Dim blnLoad, rsCat, strQuery, blnUseSourceRS
    blnLoad = True
    blnUseSourceRS = False
    If IsObject(SourceRS) Then
      If SourceRS.State <> adStateClosed Then
        Set rsCat = SourceRS
        blnUseSourceRS = True
      Else
        blnLoad = False
        m_LastError = "SourceRS is closed."
      End If
    Else
      If IsNumeric(m_ID) Then
        strQuery = "SELECT * FROM tblCategories WHERE CatPK = " & m_ID
        Set rsCat = Server.CreateObject("ADODB.RecordSet")
        rsCat.Open strQuery, m_cnnDB
      Else
        blnLoad = False
        m_LastError = "Missing ID."
      End If
    End If
    If blnLoad Then
      If Not rsCat.EOF And Not rsCat.BOF Then
        On Error Resume Next
        m_ID = rsCat("CatPK")
        m_OrgID = rsCat("OrgFK")
        m_IsActive = rsCat("IsActive")
        m_CatName = rsCat("CatName")
        m_CaseTypeID = rsCat("CaseTypeFK")
        m_CatDesc = rsCat("CatDesc")
        m_CatOrder = rsCat("CatOrder")
        m_LastUpdate = rsCat("LastUpdate")
        m_LastUpdateBy = rsCat("LastUpdateByFK")
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
      rsCat.Close
      Set rsCat = Nothing
    End If
    If blnLoad Then
      m_IsValid = True
    End If
    Load = blnLoad
  End Function

  Public Function Update() ' As Boolean
    Dim blnUpdate, rsCat
    blnUpdate = True
    Set rsCat = Server.CreateObject("ADODB.RecordSet")
    If m_IsValid Then
      rsCat.Open "SELECT * FROM tblCategories WHERE CatPK = " & m_ID, m_cnnDB, _
        adOpenKeyset, adLockOptimistic, adCmdText
    Else
      rsCat.Open "tblCategories", m_cnnDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect
      rsCat.AddNew
    End If
    With rsCat
      .Fields("IsActive") = m_IsActive
      .Fields("CaseTypeFK") = m_CaseTypeID
      .Fields("CatName") = m_CatName
      .Fields("CatDesc") = m_CatDesc
      .Fields("CatOrder") = m_CatOrder
      .Fields("LastUpdate") = m_LastUpdate
      .Fields("LastUpdateByFK") = m_LastUpdateByID
      .Update
      m_ID = .Fields("CatPK")
    End With
    rsCat.Close
    Set rsCat = Nothing
    m_IsValid = True
    Update = blnUpdate
  End Function

  Public Function Delete() ' As Boolean
    Dim blnDelete
    blnDelete = True
    m_IsActive = False
    If m_IsValid Then
      m_cnnDB.Execute "UPDATE tblCategories SET IsActive = " & m_IsActive & _
        " WHERE CatPK = " & m_ID, , adExecuteNoRecords
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
'    m_CatName = "Unknown"
  End Sub

  Private Sub Class_Terminate()
  End Sub

End Class
%>