<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: clsLanguage.asp
'  Date:     $Date: 2004/03/11 06:29:08 $
'  Version:  $Revision: 1.2 $
'  Purpose:  Class object to assist when working with Languages
' ----------------------------------------------------------------------------------

Class clsLanguage
  ' ##################
  ' Private Properties
  ' ##################
  Private m_ID, m_LangName, m_Localized
  Private m_IsRTL, m_Encoding, m_ISO639
  Private m_LastUpdate, m_LastUpdateBy, m_LastUpdateByID

  Private m_IsValid, m_IsActive, m_LastError, m_cnnDB

  ' #################
  ' Public Properties
  ' #################
  Public SourceRS

  Public Property Get ID()  ' As Long
    ID = m_ID
  End Property

  Public Property Let ID(ByRef f_ID)
    If IsNumeric(f_ID) Then
      m_ID = f_ID
    End If
  End Property

  Public Property Get LangName()  ' As String
    LangName = m_LangName
  End Property

  Public Property Let LangName(ByRef f_LangName)
    m_LangName = Left(Trim(f_LangName), 24)
  End Property

  Public Property Get Localized()  ' As String
    Localized = m_Localized
  End Property

  Public Property Let Localized(ByRef f_Localized)
    m_Localized = Left(Trim(f_Localized), 24)
  End Property

  Public Property Get IsRTL() ' As Boolean
    IsRTL = m_IsRTL
  End Property

  Public Property Let IsRTL(ByRef f_IsRTL)
    If f_IsRTL Then
      m_IsRTL = True
    Else
      m_IsRTL = False
    End If
  End Property

  Public Property Get Encoding()  ' As String
    Encoding = m_Encoding
  End Property

  Public Property Let Encoding(ByRef f_Encoding)
    m_Encoding = Left(Trim(f_Encoding), 20)
  End Property

  Public Property Get ISO639()  ' As String
    ISO639 = m_ISO639
  End Property

  Public Property Let ISO639(ByRef f_ISO639)
    m_ISO639 = Left(Trim(f_ISO639), 2)
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

    Dim blnLoad, rsLang, strQuery, blnUseSourceRS
    
    
    blnLoad = True
    blnUseSourceRS = False
    
    If IsObject(SourceRS) Then
    
      If SourceRS.State <> adStateClosed Then
        Set rsLang = SourceRS
        blnUseSourceRS = True
      Else
        blnLoad = False
        m_LastError = "SourceRS is closed."
      End If
      
    Else
    
      If IsNumeric(m_ID) And Not IsEmpty(m_ID) Then
        strQuery = "SELECT * FROM tblLanguages WHERE LangPK = " & m_ID
        Set rsLang = Server.CreateObject("ADODB.RecordSet")
        rsLang.Open strQuery, m_cnnDB
      Else
        blnLoad = False
        m_LastError = "Missing ID."
      End If
      
    End If
    
    If blnLoad Then
      If Not rsLang.EOF And Not rsLang.BOF Then
        On Error Resume Next
        m_ID = rsLang("LangPK")
        m_LangName = rsLang("LangName")
        m_Localized = rsLang("Localized")
        m_IsRTL = rsLang("IsRTL")
        m_Encoding = rsLang("Encoding")
        m_IsActive = rsLang("IsActive")
        m_LastUpdate = rsLang("LastUpdate")
        m_LastUpdateBy = rsLang("LastUpdateByFK")
        On Error Goto 0
        If Err.Number <> 0 Then
          blnLoad = False
          m_LastError = "Error retrieving data -- " & Err.Source & ": " & Err.Description
        End If
      Else
        blnLoad = False
        m_LastError = "No matching languages found."
      End If
    End If
    
    If Not blnUseSourceRS And blnLoad Then
      rsLang.Close
      Set rsLang = Nothing
    End If
    
    If blnLoad Then
      m_IsValid = True
    End If
    
    Load = blnLoad
  End Function

  Public Function Update() ' As Boolean
    Dim blnUpdate, rsLang
    blnUpdate = True
    Set rsLang = Server.CreateObject("ADODB.RecordSet")
    If m_IsValid Then
      rsLang.Open "SELECT * FROM tblLanguages WHERE LangPK = " & m_ID, m_cnnDB, _
        adOpenKeyset, adLockOptimistic, adCmdText
    Else
      rsLang.Open "tblLanguages", m_cnnDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect
      rsLang.AddNew
    End If
    With rsLang
      .Fields("LangName") = m_LangName
      .Fields("Localized") = m_Localized
      .Fields("IsRTL") = m_IsRTL
      .Fields("Encoding") = m_Encoding
      .Fields("ISO639") = m_ISO639
      .Fields("IsActive") = m_IsActive
      .Fields("LastUpdate") = m_LastUpdate
      .Fields("LastUpdateByFK") = m_LastUpdateByID
      .Update
      m_ID = .Fields("LangPK")
    End With
    rsLang.Close
    Set rsLang = Nothing
    m_IsValid = True
    Update = blnUpdate
  End Function

  Public Function Delete() ' As Boolean
    Dim blnDelete
    blnDelete = True
    m_IsActive = False
    If m_IsValid Then
      m_cnnDB.Execute "UPDATE tblLanguages SET IsActive = " & m_IsActive & _
        " WHERE LangPK = " & m_ID, , adExecuteNoRecords
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

    ' Set Default Lang Properties
    m_LangName = "Unknown"
    m_Localized = "Unknown"
    m_Encoding = "UTF-8"
    m_IsRTL = False
  End Sub

  Private Sub Class_Terminate()
  End Sub
End Class
%>