<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: clsLanguagetext.asp
'  Date:     $Date: 2004/03/17 15:54:58 $
'  Version:  $Revision: 1.2 $
'  Purpose:  Class for managing Language Text
' ----------------------------------------------------------------------------------

Class clsLanguageText
  ' ##################
  ' Private Properties
  ' ##################
  Private m_ID, m_LangID
  Private m_LangLabel, m_LangLabelID
  Private m_LangText
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

  Public Property Get LangID() ' As Long
    LangID = m_LangID
  End Property

  Public Property Let LangID(ByVal f_LangID)
    Dim oLang
    Set oLang = New clsLanguage
    oLang.ID = f_LangID
    If Not oLang.Load Then
      m_LastError = "Cannot_find_that_Language"
    End If
    m_LangID = f_LangID
  End Property

  Public Property Get LangLabelID()  ' As Integer
    LangLabelID = m_LangLabelID
  End Property

  Public Property Let LangLabelID(ByRef f_LangLabelID)
    If IsNumeric(f_LangLabelID) Then
      m_LangLabelID = CInt(f_LangLabelID)
    Else
      m_LastError = "Language_ID_Must_Be_Numeric"
    End If
  End Property


  Public Property Get LangLabel()  ' As String

    Dim objLangLabel

    If Not IsObject(m_LangLabel) Then

      Set objLangLabel = New clsLanguageLabel

      If IsNumeric(m_LangLabelID) And Not IsEmpty(m_LangLabelID) Then

        objLangLabel.ID = m_LangLabelID

        If Not objLangLabel.Load Then
    		   objLangLabel.LangLabel = ""
        Else
           ' Label loaded
        End If

      Else

         ' Invalid LangLabelID
         objLangLabel.LangLabel = ""

      End If
      
      Set m_LangLabel = objLangLabel
      Set LangLabel = m_LangLabel

      Set objLangLabel = Nothing
    
    Else
    
		  ' Label already loaded

    End If

  End Property

  Public Property Let LangLabel(ByRef f_LangLabel)
    m_LangLabel = Trim(f_LangLabel)
  End Property

  Public Property Get LangText()  ' As String
    LangText = m_LangText
  End Property

  Public Property Let LangText(ByRef f_LangText)
    m_LangText = Trim(f_LangText)
  End Property

  Public Property Get AddDate()  ' As Date
    AddDate = m_AddDate
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

    Dim blnLoad, rsLangText, strQuery, blnUseSourceRS
    blnLoad = True
    blnUseSourceRS = False
    If IsObject(SourceRS) Then
      If SourceRS.State <> adStateClosed Then
        Set rsLangText = SourceRS
        blnUseSourceRS = True
      Else
        blnLoad = False
        m_LastError = "SourceRS is closed."
      End If
    Else
      If IsNumeric(m_ID) Then
        strQuery = "SELECT * FROM tblLanguageTexts WHERE LangTextPK = " & m_ID

        'Try this on the webpage: like '*name*' or langtext like '*name*'; "
        'This query should work with Access.  
        'Probably needs to be changed for SQL.
        Set rsLangText = Server.CreateObject("ADODB.RecordSet")
        rsLangText.Open strQuery, m_cnnDB
      Else
        blnLoad = False
        m_LastError = "Missing ID."
      End If
    End If
    If blnLoad Then
      If Not rsLangText.EOF And Not rsLangText.BOF Then
        On Error Resume Next
        m_ID = rsLangText("LangTextPK")
        m_LangID = rsLangText("LangFK")
        m_LangLabelID = rsLangText("LangLabelFK")
        m_LangText = rsLangText("LangText")
        m_LastUpdate = rsLangText("LastUpdate")
        m_LastUpdateByID = rsLangText("LastUpdateByFK")
        On Error Goto 0
        If Err.Number <> 0 Then
          blnLoad = False
          m_LastError = "Error retrieving data -- " & Err.Source & ": " & Err.Description
        End If
      Else
        blnLoad = False
        m_LastError = "No matching LangTexts found."
      End If
    End If
    If Not blnUseSourceRS Then
      rsLangText.Close
      Set rsLangText = Nothing
    End If
    If blnLoad Then
      m_IsValid = True
    End If
    Load = blnLoad
  End Function

  Public Function Update() ' As Boolean
    Dim blnUpdate, rsLangText
    blnUpdate = True
    Set rsLangText = Server.CreateObject("ADODB.RecordSet")
    If m_IsValid Then
      rsLangText.Open "SELECT * FROM tblLanguageTexts WHERE LangTextPK = " & m_ID, m_cnnDB, _
        adOpenKeyset, adLockOptimistic, adCmdText
    Else
      'Need to add code here for insertion of the LangLabel if new.
      rsLangText.Open "tblLanguageTexts", m_cnnDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect
      rsLangText.AddNew
    End If
    With rsLangText
      .Fields("LangFK") = m_LangID
      .Fields("LangLabelFK") = m_LangLabelID
      .Fields("LangText") = m_LangText
      .Fields("LastUpdate") = m_LastUpdate
      .Fields("LastUpdateByFK") = m_LastUpdateByID
      .Update
      m_ID = .Fields("LangTextPK")
    End With
    rsLangText.Close
    Set rsLangText = Nothing
    m_IsValid = True
    Update = blnUpdate
  End Function

  Public Function Delete() ' As Boolean
    Dim blnDelete
    blnDelete = True
    If m_IsValid Then
      m_cnnDB.Execute "DELETE FROM tblLanguageTexts" & _
        " WHERE LangTextPK = " & m_ID, , adExecuteNoRecords
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