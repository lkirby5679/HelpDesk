<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: clsLanguageLabel.asp
'  Date:     $Date: 2004/03/17 00:06:45 $
'  Version:  $Revision: 1.1 $
'  Purpose:  Class object to assist when working with Contacts
' ----------------------------------------------------------------------------------

Class clsLanguageLabel

  ' ##################
  ' Private Properties
  ' ##################
  
  Private m_ID, m_LangLabel
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
    Else
      m_LastError = "Text_ID_Must_Be_Numeric"
    End If
  End Property

  Public Property Get LangLabel()  ' As String
    LangLabel = m_LangLabel
  End Property

  Public Property Let LangLabel(ByRef f_LangLabel)
    m_LangLabel = Trim(f_LangLabel)
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

    Dim blnLoad, rsLangLabel, strQuery, blnUseSourceRS
    blnLoad = True
    blnUseSourceRS = False
    If IsObject(SourceRS) Then
      If SourceRS.State <> adStateClosed Then
        Set rsLangLabel = SourceRS
        blnUseSourceRS = True
      Else
        blnLoad = False
        m_LastError = "SourceRS is closed."
      End If
    Else
      If IsNumeric(m_ID) Then
        strQuery = "SELECT * FROM tblLanguageLabels WHERE LangLabelPK = " & m_ID
        'Try this on the webpage: like '*name*' or langtext like '*name*'; "
        'This query should work with Access.  
        'Probably needs to be changed for SQL.
        Set rsLangLabel = Server.CreateObject("ADODB.RecordSet")
        rsLangLabel.Open strQuery, m_cnnDB
      Else
        blnLoad = False
        m_LastError = "Missing ID."
      End If
    End If
    If blnLoad Then
      If Not rsLangLabel.EOF And Not rsLangLabel.BOF Then
        On Error Resume Next
        m_ID = rsLangLabel("LangLabelPK")
        m_LangLabel = rsLangLabel("LangLabel")
        m_LastUpdate = rsLangLabel("LastUpdate")
        m_LastUpdateByID = rsLangLabel("LastUpdateByFK")
        On Error Goto 0
        If Err.Number <> 0 Then
          blnLoad = False
          m_LastError = "Error retrieving data -- " & Err.Source & ": " & Err.Description
        End If
      Else
        blnLoad = False
        m_LastError = "No matching LangLabels found."
      End If
    End If
    If Not blnUseSourceRS Then
      rsLangLabel.Close
      Set rsLangLabel = Nothing
    End If
    If blnLoad Then
      m_IsValid = True
    End If
    Load = blnLoad
  End Function

  Public Function Update() ' As Boolean
    Dim blnUpdate, rsLangLabel
    blnUpdate = True
    Set rsLangLabel = Server.CreateObject("ADODB.RecordSet")
    If m_IsValid Then
      rsLangLabel.Open "SELECT * FROM tblLanguageLabels WHERE LangLabelPK = " & m_ID, m_cnnDB, _
        adOpenKeyset, adLockOptimistic, adCmdText
    Else
      'Need to add code here for insertion of the LangLabel if new.
      rsLangLabel.Open "tblLanguageLabels", m_cnnDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect
      rsLangLabel.AddNew
    End If
    With rsLangLabel
      .Fields("LangLabel") = m_LangLabel
      .Fields("LastUpdate") = m_LastUpdate
      .Fields("LastUpdateByFK") = m_LastUpdateByID
      .Update
      m_ID = .Fields("LangLabelPK")
    End With
    rsLangLabel.Close
    Set rsLangLabel = Nothing
    m_IsValid = True
    Update = blnUpdate
  End Function

  Public Function Delete() ' As Boolean
    Dim blnDelete
    blnDelete = True
    If m_IsValid Then
      m_cnnDB.Execute "DELETE FROM tblLanguageLabels" & _
        " WHERE LangLabelPK = " & m_ID, , adExecuteNoRecords
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

  End Sub

  Private Sub Class_Terminate()

  End Sub

End Class
%>