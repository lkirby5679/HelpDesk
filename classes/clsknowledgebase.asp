<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: clsKnowledgebase.asp
'  Date:     $Date: 2004/03/11 06:29:08 $
'  Version:  $Revision: 1.2 $
'  Purpose:  Class object to assist when working with Knowledgebase Records
' ----------------------------------------------------------------------------------

Class clsKnowledgebase

  ' ##################
  ' Private Properties
  ' ##################
  
  Private m_ID, m_Issue, m_Cause, m_Resolution, m_IsActive, m_EnteredDate
  Private m_LastUpdate, m_LastUpdateBy, m_LastUpdateByID, m_EnteredBy, m_EnteredByID
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

  Public Property Get Issue() ' As String
    Issue = m_Issue
  End Property

  Public Property Let Issue(ByRef f_Issue)
    m_Issue = Trim(f_Issue)
  End Property

  Public Property Get Cause() ' As String
    Cause = m_Cause
  End Property

  Public Property Let Cause(ByRef f_Cause)
    m_Cause = Trim(f_Cause)
  End Property

  Public Property Get Resolution() ' As String
    Resolution = m_Resolution
  End Property

  Public Property Let Resolution(ByRef f_Resolution)
    m_Resolution = Trim(f_Resolution)
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

  Public Property Get IsActive() ' As Boolean
    IsActive = m_IsActive
  End Property

  Public Property Let IsActive(f_IsActive) ' As Boolean
    If IsNumeric(f_IsActive) Then
  		m_IsActive=f_IsActive
    End If
    IsActive = m_IsActive
  End Property

  Public Property Get EnteredBy() ' As clsContact
  
    Dim objContact, blnTemp

    If Not IsObject(m_EnteredBy) Then
    
      Set objContact = New clsContact
      If IsNumeric(m_EnteredByID) Then
        objContact.ID = m_EnteredByID
        blnTemp = objContact.Load
      End If
      Set m_EnteredBy = objContact
      
    End If

    Set EnteredBy = m_EnteredBy
    
  End Property

  Public Property Get EnteredByID() ' As Long
    EnteredByID = m_EnteredByID
  End Property

  Public Property Let EnteredByID(f_EnteredByID) ' As Long
    m_EnteredByID = f_EnteredByID
  End Property

  Public Property Get EnteredDate() ' As Date
    EnteredDate = m_EnteredDate
  End Property

  Public Property Let EnteredDate(ByRef f_dtEnteredDate)
    If Not IsDate(f_dtEnteredDate) Then
      ' #######
      ' Print an error and die
    End If
    m_EnteredDate = f_dtEnteredDate
  End Property

  Public Property Get LastUpdateBy() ' As clsContact
  
    Dim objContact, blnTemp

    If Not IsObject(m_LastUpdateBy) Then
    
      Set objContact = New clsContact
      If IsNumeric(m_LastUpdateByID) Then
        objContact.ID = m_LastUpdateByID
        blnTemp = objContact.Load
      End If
      Set m_LastUpdateBy = objContact
      
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

    Dim blnLoad, rstKnowledgebase, strQuery, blnUseSourceRS

    blnLoad = True
    blnUseSourceRS = False

    If IsObject(SourceRS) Then
      If SourceRS.State <> adStateClosed Then
        Set rstKnowledgebase = SourceRS
        blnUseSourceRS = True
      Else
        blnLoad = False
        m_LastError = "SourceRS is closed."
      End If
    Else
      If IsNumeric(m_ID) And Not IsEmpty(m_ID) Then
        Set rstKnowledgebase = Server.CreateObject("ADODB.RecordSet")
        rstKnowledgebase.Open "SELECT * FROM tblKnowledgebase WHERE KnowledgebasePK=" & m_ID, m_cnnDB
      Else
        blnLoad = False
        m_LastError = "Missing ID."
      End If
    End If

    If blnLoad Then

      If Not rstKnowledgebase.EOF And Not rstKnowledgebase.BOF Then
        On Error Resume Next
          m_ID = rstKnowledgebase("KnowledgebasePK")
          m_Issue = rstKnowledgebase("Issue")
          m_Cause = rstKnowledgebase("Cause")
          m_Resolution = rstKnowledgebase("Resolution")
          m_EnteredByID = rstKnowledgebase("EnteredByFK")
          m_EnteredDate = rstKnowledgebase("EnteredDate")
          m_IsActive = rstKnowledgebase("IsActive")
          m_LastUpdate = rstKnowledgebase("LastUpdate")
          m_LastUpdateByID = rstKnowledgebase("LastUpdateByFK")
        On Error Goto 0
          If Err.Number <> 0 Then
            blnLoad = False
            m_LastError = "Error retrieving data -- " & Err.Source & ": " & Err.Description
          End If

      Else
        blnLoad = False
        m_LastError = "No matching Knowledgebases found."

      End If

    End If

    If Not blnUseSourceRS And IsObject(rstKnowledgebase)Then
      rstKnowledgebase.Close
      Set rstKnowledgebase = Nothing
    End If

    If blnLoad Then
      m_IsValid = True
    End If

    Load = blnLoad

  End Function

  Public Function Update() ' As Boolean
    Dim blnUpdate, rstKnowledgebase, strSQL
    
    blnUpdate = True
    
    Set rstKnowledgebase = Server.CreateObject("ADODB.RecordSet")
    
    If m_IsValid Then
      rstKnowledgebase.Open "SELECT * FROM tblKnowledgebase WHERE KnowledgebasePK=" & m_ID, m_cnnDB, adOpenKeyset, adLockOptimistic, adCmdText

    Else
      rstKnowledgebase.Open "tblKnowledgebase", m_cnnDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect
      rstKnowledgebase.AddNew
    End If

  	With rstKnowledgebase
  	  .Fields("Issue") = m_Issue
  	  .Fields("Cause") = m_Cause
  	  .Fields("Resolution") = m_Resolution
  	  .Fields("EnteredByFK") = m_EnteredByID
  	  .Fields("EnteredDate") = m_EnteredDate
  	  .Fields("IsActive") = m_IsActive
  	  .Fields("LastUpdate") = m_LastUpdate
  	  .Fields("LastUpdateByFK") = m_LastUpdateByID
  	  .Update
  	End With
  	
    m_ID = rstKnowledgebase.Fields("KnowledgebasePK")

    rstKnowledgebase.Close
    Set rstKnowledgebase = Nothing

    m_IsValid = True

    Update = blnUpdate

  End Function

  ' ###############
  ' Private Methods
  ' ###############

  Private Sub Class_Initialize()

    If IsObject(cnnDB) Then
      Set m_cnnDB = cnnDB
    End If

    m_IsValid = False
    m_EnteredDate = Now()

  End Sub

  Private Sub Class_Terminate()

  End Sub

End Class
%>