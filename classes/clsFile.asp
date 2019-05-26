<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: clsFile.asp
'  Date:     $Date: 2004/03/11 06:29:08 $
'  Version:  $Revision: 1.2 $
'  Purpose:  Class object to assist when working with Attachments
' ----------------------------------------------------------------------------------

Class clsFile
  ' ##################
  ' Private Properties
  ' ##################
  Private m_ID
  Private m_CaseID, m_IsFile, m_FileName, m_FileSize, m_FileLocation
  Private m_ContentType, m_UploadDate
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

  Public Property Get IsFile() ' As Boolean
    IsFile = m_IsFile
  End Property

  Public Property Let IsFile(ByRef f_IsFile)
    If f_IsFile Then
      m_IsFile = True
    Else
      m_IsFile = False
    End If
  End Property

  Public Property Get FileName()  ' As String
    FileName = m_FileName
  End Property

  Public Property Let FileName(ByRef f_FileName)
    m_FileName = Left(Trim(f_FileName), 255)
  End Property

  Public Property Get FileLocation()  ' As String
    FileLocation = m_FileLocation
  End Property

  Public Property Let FileLocation(ByRef f_FileLocation)
    m_FileLocation = Left(Trim(f_FileLocation), 255)
  End Property

  Public Property Get FileSize()  ' As Long
    FileSize = m_FileSize
  End Property

  Public Property Get ContentType()  ' As String
    Note = m_Note
  End Property

  Public Property Let ContentType(ByRef f_ContentType)
    m_ContentType = Left(Trim(f_ContentType), 50)
  End Property

  Public Property Get UploadDate()  ' As Date
    AddDate = m_AddDate
  End Property

  Public Property Get LastUpdate() ' As Date
    LastUpdate = m_LastUpdate
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

    Dim blnLoad, rsFile, strQuery, blnUseSourceRS
    blnLoad = True
    blnUseSourceRS = False
    If IsObject(SourceRS) Then
      If SourceRS.State <> adStateClosed Then
        Set rsFile = SourceRS
        blnUseSourceRS = True
      Else
        blnLoad = False
        m_LastError = "SourceRS is closed."
      End If
    Else
      If IsNumeric(m_ID) Then
        strQuery = "SELECT FilePK, CaseFK, IsFile, FileName, FileSize, " & _
          "ContentType, UploadDate, LastUpdate, LastUpDateByFK FROM tblFiles WHERE FilePK = " & m_ID
        Set rsFile = Server.CreateObject("ADODB.RecordSet")
        rsFile.Open strQuery, m_cnnDB
      Else
        blnLoad = False
        m_LastError = "Missing ID."
      End If
    End If
    If blnLoad Then
      If Not rsFile.EOF And Not rsFile.BOF Then
        On Error Resume Next
        m_ID = rsFile("FilePK")
        m_IsFile = rsFile("IsFile")
        m_FileName = rsFile("FileName")
        m_FileSize = rsFile("FileSize")
        m_ContentType = rsFile("ContentType")
        m_FileLocation = rsFile("FileLocation")
        m_UploadDate = rsFile("UploadDate")
        m_LastUpdate = rsFile("LastUpdate")
        m_LastUpdateByID = rsFile("LastUpdateByFK")
        On Error Goto 0
        If Err.Number <> 0 Then
          blnLoad = False
          m_LastError = "Error retrieving data -- " & Err.Source & ": " & Err.Description
        End If
      Else
        blnLoad = False
        m_LastError = "No matching files found."
      End If
    End If
    If Not blnUseSourceRS Then
      rsFile.Close
      Set rsFile = Nothing
    End If
    If blnLoad Then
      m_IsValid = True
    End If
    Load = blnLoad
  End Function

  Public Function Update() ' As Boolean
    Dim blnUpdate, rsFile
    blnUpdate = True
    Set rsFile = Server.CreateObject("ADODB.RecordSet")
    If m_IsValid Then
      rsFile.Open "SELECT FilePK, CaseFK, IsFile, FileName, FileSize, " & _
        "ContentType, UploadDate, LastUpdate, LastUpDateByFK FROM tblFiles WHERE FilePK = " & m_ID, m_cnnDB, _
        adOpenKeyset, adLockOptimistic, adCmdText
    Else
      rsFile.Open "tblFiles", m_cnnDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect
      rsFile.AddNew
    End If
    With rsFile
      .Fields("CaseFK") = m_CaseID
      .Fields("IsFile") = m_IsFile
      .Fields("FileName") = m_FileName
      .Fields("FileSize") = m_FileSize
      .Fields("ContentType") = m_ContentType
      .Fields("FileLocation") = m_FileLocation
      .Fields("UploadDate") = m_UploadDate
      .Fields("LastUpdate") = m_LastUpdate
      .Fields("LastUpdateByFK") = m_LastUpdateByID
      .Update
      m_ID = .Fields("FilePK")
    End With
    rsFile.Close
    Set rsFile = Nothing
    m_IsValid = True
    Update = blnUpdate
  End Function

  Public Function Delete() ' As Boolean
    Dim blnDelete
    blnDelete = True
    If m_IsValid Then
      m_cnnDB.Execute "DELETE FROM tblFiles" & _
        " WHERE FileFK = " & m_ID, , adExecuteNoRecords
      If m_IsFile Then
        ' ##### Remove files on the file system
      End If
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
    m_UploadDate = Now
    m_IsFile = False
  End Sub

  Private Sub Class_Terminate()
  End Sub

End Class
%>