<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: clsListItem.asp
'  Date:     $Date: 2004/03/11 06:29:08 $
'  Version:  $Revision: 1.2 $
'  Purpose:  Class object to assist when working with List Items
' ----------------------------------------------------------------------------------

Class clsListItem
  ' ##################
  ' Private Properties
  ' ##################
  Private m_ID
  Private m_ParentListItemID, m_ParentListItem
  Private m_ItemOrder, m_IsActive, m_ItemName
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

  Public Property Get ParentListItem() ' As clsParentListItem
    If Not IsObject(m_ParentListItem) Then
      Dim oParentListItem, blnTemp
      Set oParentListItem = New clsListItem
      If IsNumeric(m_ParentListItemID) Then
        oParentListItem.ID = m_ParentListItemID
        blnTemp = oParentListItem.Load  ' Ignore Errors
      End If
      Set m_ParentListItem = oParentListItem
    End If
    Set ParentListItem = m_ParentListItem
  End Property

  Public Property Get ParentListItemID() ' As Long
    ParentListItemID = m_ParentListItemID    
  End Property

  Public Property Let ParentListItemID(ByVal f_ParentListItemID)
    Dim oParentListItem
    Set oParentListItem = New clsListItem
    oParentListItem.ID = f_ParentListItemID
    If Not oParentListItem.Load Then
      ' ########
      ' Raise Error
    End If
    m_ParentListItemID = f_ParentListItemID
    Set m_ParentListItem = oParentListItem
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

  Public Property Get ItemOrder()  ' As Long
    ItemOrder = m_ItemOrder
  End Property

  Public Property Let ItemOrder(ByRef f_ItemOrder)
    If IsNumeric(f_ItemOrder) Then
      m_ItemOrder = f_ItemOrder
    Else
      m_ItemOrder = 0
    End If
  End Property

  Public Property Get ItemName()  ' As String
    ItemName = m_ItemName
  End Property

  Public Property Let ItemName(ByRef f_ItemName)
    m_ItemName = Left(Trim(f_ItemName), 50)
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

    Dim blnLoad, rsList, strQuery, blnUseSourceRS
    blnLoad = True
    blnUseSourceRS = False
    If IsObject(SourceRS) Then
      If SourceRS.State <> adStateClosed Then
        Set rsList = SourceRS
        blnUseSourceRS = True
      Else
        blnLoad = False
        m_LastError = "SourceRS is closed."
      End If
    Else
      If IsNumeric(m_ID) Then
        strQuery = "SELECT * FROM tblLists WHERE ListItemPK = " & m_ID
        Set rsList = Server.CreateObject("ADODB.RecordSet")
        rsList.Open strQuery, m_cnnDB
      Else
        blnLoad = False
        m_LastError = "Missing ID."
      End If
    End If
    If blnLoad Then
      If Not rsList.EOF And Not rsList.BOF Then
        On Error Resume Next
        m_ID = rsList("ListItemPK")
        m_ParentListItemID = rsList("ParentListItemFK")
        m_ItemOrder = rsList("ItemOrder")
        m_IsActive = rsList("IsActive")
        m_ItemName = rsList("ItemName")
        m_LastUpdate = rsList("LastUpdate")
        m_LastUpdateByID = rsList("LastUpdateByFK")
        On Error Goto 0
        If Err.Number <> 0 Then
          blnLoad = False
          m_LastError = "Error retrieving data -- " & Err.Source & ": " & Err.Description
        End If
      Else
        blnLoad = False
        m_LastError = "No matching items found."
      End If
    End If
    If Not blnUseSourceRS Then
      rsList.Close
      Set rsList = Nothing
    End If
    If blnLoad Then
      m_IsValid = True
    End If
    Load = blnLoad
  End Function

  Public Function Update() ' As Boolean
    Dim blnUpdate, rsList
    blnUpdate = True
    Set rsList = Server.CreateObject("ADODB.RecordSet")
    If m_IsValid Then
      rsList.Open "SELECT * FROM tblLists WHERE ListItemPK = " & m_ID, m_cnnDB, _
        adOpenKeyset, adLockOptimistic, adCmdText
    Else
      rsList.Open "tblLists", m_cnnDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect
      rsList.AddNew
    End If
    With rsList
      .Fields("ParentListItemFK") = m_ParentListItemID
      .Fields("ItemOrder") = m_ItemOrder
      .Fields("IsActive") = m_IsActive
      .Fields("ItemName") = m_ItemName
      .Fields("LastUpdate") = m_LastUpdate
      .Fields("LastUpdateByFK") = m_LastUpdateByID
      .Update
      m_ID = .Fields("ListItemPK")
    End With
    rsList.Close
    Set rsList = Nothing
    m_IsValid = True
    Update = blnUpdate
  End Function

  Public Function Delete() ' As Boolean
    Dim blnDelete
    blnDelete = True
    If m_IsValid Then
      m_cnnDB.Execute "UPDATE tblLists SET IsActive = " & m_IsActive & _
        " WHERE ListItemPK = " & m_ID, , adExecuteNoRecords
    End If
    Delete = blnDelete
    m_IsValid = False
  End Function

  ' ###############
  ' Private Methods
  ' ###############
  Private Sub Class_Initialize()
    If IsObject(cnnDB) Then
      Set m_cnnDB = cnnDB
    End If
    m_IsValid = False

'    ' Set Default Properties
'    m_ItemOrder = 0
'    m_IsActive = True
  End Sub

  Private Sub Class_Terminate()
  End Sub

End Class
%>