<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: clsNote.asp
'  Date:     $Date: 2004/03/11 06:29:08 $
'  Version:  $Revision: 1.2 $
'  Purpose:  Class object to assist when working with Notes and Cases
' ----------------------------------------------------------------------------------

Class clsNote
  ' ##################
  ' Private Properties
  ' ##################
  Private m_ID, m_IsActive
  Private m_CaseID
  Private m_Note, m_AddDate, m_IsPrivate, m_MinutesSpent
  Private m_OwnerID, m_Owner, m_BillTypeID, m_BillType
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

  Public Property Get CaseID() ' As Long
    CaseID = m_CaseID
  End Property

  Public Property Let CaseID(ByVal f_CaseID)
    Dim oCase
    Set oCase = New clsCase
    oCase.ID = f_CaseID
    If Not oCase.Load Then
      ' ########
      ' Raise Error
    End If
    m_CaseID = f_CaseID
  End Property


  Public Property Get Note()  ' As String
    Note = m_Note
  End Property

  Public Property Let Note(ByRef f_Note)
    m_Note = Trim(f_Note)
  End Property

  Public Property Get AddDate()  ' As Date
    AddDate = m_AddDate
  End Property

  Public Property Let AddDate(ByRef f_dtAddDate)
    If Not IsDate(f_dtAddDate) Then
      ' #######
      ' Print an error and die
    End If
    m_AddDate = f_dtAddDate
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

  Public Property Get IsPrivate() ' As Boolean
    IsPrivate = m_IsPrivate
  End Property

  Public Property Let IsPrivate(ByRef f_IsPrivate)
    If f_IsPrivate Then
      m_IsPrivate = True
    Else
      m_IsPrivate = False
    End If
  End Property

  Public Property Get MinutesSpent()  ' As Long
    MinutesSpent = m_MinutesSpent
  End Property

  Public Property Let MinutesSpent(ByRef f_MinutesSpent)
    If IsNumeric(f_MinutesSpent) Then
      m_MinutesSpent = f_MinutesSpent
    Else
      m_MinutesSpent = 0
    End If
  End Property

  Public Property Get Owner()  ' As clsContact
    If Not IsObject(m_Owner) Then
      Dim oOwner, blnTemp
      If IsNumeric(m_OwnerID) Then
        Set oOwner = New clsContact
        oOwner.ID = m_OwnerID
      End If
      blnTemp = oOwner.Load
      Set m_Owner = oOwner
    End If
    Set Owner = m_Owner
  End Property

  Public Property Get OwnerID()  'As Long
    OwnerID = m_OwnerID
  End Property

  Public Property Let OwnerID(ByRef f_OwnerID)
    Dim oOwner
    Set oOwner = New clsContact
    oOwner.ID = f_OwnerID
    If Not oOwner.Load Then
      ' ######
      ' Print LastError
    End If
    m_OwnerID = f_OwnerID
    Set m_Owner = oOwner
  End Property

  Public Property Get BillType() ' As clsListItem

    Dim objBillType

    If Not IsObject(m_BillType) Then
    
      Set objBillType = New clsListItem
      If IsNumeric(m_BillTypeID) And Not IsEmpty(m_BillTypeID) Then
        objBillType.ID = m_BillTypeID
        If Not oBillType.Load Then
		  ' List Item not found
        Else
		  ' List Item found
		  Set m_BillType = objBillType
        End If
      Else
        ' Invalid List Item ID
      End If
      Set objBillType = Nothing
      
    Else
    
	  ' BillType already loaded
      
    End If
    
    Set BillType = m_BillType

  End Property

  Public Property Get BillTypeID() ' As Long
    BillTypeID = m_BillTypeID
  End Property

  Public Property Let BillTypeID(ByRef f_BillTypeID)

    Dim objBillType
    
    If IsNumeric(f_BillTypeID) And Not IsEmpty(f_BillTypeID) Then
    
      Set objBillType = New clsListItem

      objBillType.ID = f_BillTypeID

      If Not objBillType.Load Then
        ' Item not found, Raise Error
      Else
         m_BillTypeID = f_BillTypeID
         Set m_BillType = objBillType
      End If
      
    Else
    
      ' Invalid ID
    
    End If
    
    Set objBillType = Nothing
    
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

    Dim blnLoad, rsNote, strQuery, blnUseSourceRS
    blnLoad = True
    blnUseSourceRS = False
    If IsObject(SourceRS) Then
      If SourceRS.State <> adStateClosed Then
        Set rsNote = SourceRS
        blnUseSourceRS = True
      Else
        blnLoad = False
        m_LastError = "SourceRS is closed."
      End If
    Else
      If IsNumeric(m_ID) Then
        strQuery = "SELECT * FROM tblNotes WHERE NotePK = " & m_ID
        Set rsNote = Server.CreateObject("ADODB.RecordSet")
        rsNote.Open strQuery, m_cnnDB
      Else
        blnLoad = False
        m_LastError = "Missing ID."
      End If
    End If
    If blnLoad Then
      If Not rsNote.EOF And Not rsNote.BOF Then
        On Error Resume Next
        m_ID = rsNote("NotePK")
        m_CaseID = rsNote("CaseFK")
        m_Note = rsNote("Note")
        m_AddDate = rsNote("AddDate")
        m_IsPrivate = rsNote("IsPrivate")
        m_MinutesSpent = rsNote("MinutesSpent")
        m_OwnerID = rsNote("OwnerFK")
        m_BillTypeID = rsNote("BillTypeFK")
        m_LastUpdate = rsNote("LastUpdate")
        m_LastUpdateByID = rsNote("LastUpdateByFK")
        On Error Goto 0
        If Err.Number <> 0 Then
          blnLoad = False
          m_LastError = "Error retrieving data -- " & Err.Source & ": " & Err.Description
        End If
      Else
        blnLoad = False
        m_LastError = "No matching notes found."
      End If
    End If
    If Not blnUseSourceRS Then
      rsNote.Close
      Set rsNote = Nothing
    End If
    If blnLoad Then
      m_IsValid = True
    End If
    Load = blnLoad
  End Function

  Public Function Update() ' As Boolean
    Dim blnUpdate, rsNote
    blnUpdate = True
    Set rsNote = Server.CreateObject("ADODB.RecordSet")
    If m_IsValid Then
      rsNote.Open "SELECT * FROM tblNotes WHERE NotePK = " & m_ID, m_cnnDB, _
        adOpenKeyset, adLockOptimistic, adCmdText
    Else
      rsNote.Open "tblNotes", m_cnnDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect
      rsNote.AddNew
    End If
    With rsNote
      .Fields("CaseFK") = m_CaseID
      .Fields("Note") = m_Note
      .Fields("AddDate") = m_AddDate
      .Fields("IsPrivate") = m_IsPrivate
      .Fields("MinutesSpent") = m_MinutesSpent
      .Fields("OwnerFK") = m_OwnerID
      .Fields("BillTypeFK") = m_BillTypeID
      .Fields("LastUpdate") = m_LastUpdate
      .Fields("LastUpdateByFK") = m_LastUpdateByID
      .Update
      m_ID = .Fields("NotePK")
    End With
    rsNote.Close
    Set rsNote = Nothing
    m_IsValid = True
    Update = blnUpdate
  End Function

  Public Function Delete() ' As Boolean
    Dim blnDelete
    blnDelete = True
    If m_IsValid Then
      m_cnnDB.Execute "DELETE FROM tblNotes" & _
        " WHERE NotePK = " & m_ID, , adExecuteNoRecords
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

    ' Set Default Properties
'    m_AddDate = Now
'    m_IsPrivate = False
  End Sub

  Private Sub Class_Terminate()
  End Sub

End Class
%>