<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: clsContact.asp
'  Date:     $Date: 2004/03/11 06:29:08 $
'  Version:  $Revision: 1.2 $
'  Purpose:  Class object to assist when working with Contacts
' ----------------------------------------------------------------------------------

Class clsContact

  ' ##################
  ' Private Properties
  ' ##################

  Private m_ID, m_Org, m_OrgID, m_FName, m_LName, m_Dept, m_DeptID
  Private m_ContactType, m_ContactTypeID, m_Lang, m_LangID
  Private m_IsActive, m_UserName, m_Password
  Private m_IOStatus, m_IOStatusID, m_IOStatusDate, m_IOStatusText
  Private m_TZOffset, m_UserPermMask, m_Role, m_RoleID
  Private m_OfficePhone, m_HomePhone, m_MobilePhone, m_JobTitle
  Private m_JobFunction, m_JobFunctionID
  Private m_Resume, m_Email, m_OfficeLocation, m_Notes
  Private m_Created, m_LastAccess
  Private m_LastUpdate, m_LastUpdateBy, m_LastUpdateByID
  Private m_PhotoFile, m_PhotoFileID, m_PagerEmail

  Private m_IsValid, m_LastError, m_cnnDB

  ' #################
  ' Public Properties
  ' #################

  Public SourceRS ' As ADODB.RecordSet

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
        blnTemp = oOrg.Load ' Ignore errors
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
      ' #######
      ' Raise Error
    End If
    m_OrgID = f_OrgID
    Set m_Org = oOrg
  End Property

  Public Property Get FName() ' As String
    FName = m_FName
  End Property

  Public Property Let FName(ByRef f_fname)
    m_FName = Left(Trim(f_fname), 16)
  End Property

  Public Property Get LName() ' As String
    LName = m_LName
  End Property

  Public Property Let LName(ByRef f_lname)
    m_LName = Left(Trim(f_lname), 32)
  End Property

  Public Property Get Dept() ' As clsDepartment

    Dim oDept

    If Not IsObject(m_Dept) Then

      Set oDept = New clsDepartment

      If IsNumeric(m_DeptID) And Not IsEmpty(m_DeptID) Then

        oDept.ID = m_DeptID
        If Not oDept.Load Then
           ' Department not found
		   oDept.DeptName = ""
        Else
           ' Department loaded
        End If

      Else

         ' Invalid DeptID
         oDept.DeptName = ""

      End If
      
      Set m_Dept = oDept
      Set Dept = m_Dept

      Set oDept = Nothing
    
    Else
    
		' Dept already loaded

    End If

  End Property

  Public Property Get DeptID() ' As Long
    DeptID = m_DeptID
  End Property

  Public Property Let DeptID(ByRef f_DeptID)
    Dim oDept
    
    If IsNumeric(f_DeptID) And Not IsEmpty(f_DeptID) Then 
       Set oDept = New clsDepartment
       oDept.ID = f_DeptID
       If Not oDept.Load Then
         ' ######
         ' Raise Error
       Else
          m_DeptID = f_DeptID
          Set m_Dept = oDept
       End If
       Set oDept = Nothing
    Else
       ' Invalid DeptID
    End If
    
  End Property

  Public Property Get ContactType() ' As clsListItem
    If IsObject(m_ContactType) Then
      Set ContactType = m_ContactType
    Else
      ' ######
      ' Create and return ContactType
    End If
  End Property

  Public Property Get ContactTypeID() ' As clsListItem
    ContactTypeID = m_ContactTypeID
  End Property

  Public Property Let ContactTypeID(ByRef f_ContactTypeID)
    If Not IsNumeric(f_ContactTypeID) Then
      ' ######
      ' Print an error message and die
    End If

    ' ######
    ' Verify against tblLists

    m_ContactTypeID = f_ContactTypeID
    Set m_ContactType = Nothing
  End Property

  Public Property Get Lang()  ' As clsLang
    If Not IsObject(m_Lang) Then
      Dim oLang, blnTemp
      Set oLang = New clsLanguage
      If IsNumeric(m_LangID) Then
        oLang.ID = m_LangID
        blnTemp = oLang.Load
      End If
      Set m_Lang = oLang
    End If
    Set Lang = m_Lang
  End Property

  Public Property Get LangID() ' As Long
    LangID = m_LangID
  End Property

  Public Property Let LangID(ByRef f_LangID)
    Dim oLang
    Set oLang = New clsLanguage
    oLang.ID = f_LangID
    If Not oLang.Load Then
      ' ######
      ' Raise Error
    End If
    m_LangID = f_LangID
    Set m_Lang = oLang
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

  Public Property Get UserName() ' As String
    UserName = m_UserName
  End Property

  Public Property Let UserName(ByVal f_UserName)
    f_UserName = Left(Trim(f_UserName), 24)
    If Len(f_UserName) = 0 Then
      ' #######
      ' Print error and die
    End If
    ' #######
    ' Check for duplicate username?
    m_UserName = f_UserName
  End Property

  Public Property Get Password() ' As String
    Password = m_Password
  End Property

  Public Property Let Password(ByRef f_Password)
    ' ########
    ' Encrypt password using MD5

    m_Password = f_Password
  End Property

  Public Property Get IOStatus() ' As clsListItem
    If IsObject(m_IOStatus) Then
      Set IOStatus = m_IOStatus
    Else
      ' ######
      ' Create IOstatus from tblLists
    End If
  End Property

  Public Property Get IOStatusID() ' As Long
    IOStatusID = m_IOStatusID
  End Property

  Public Property Let IOStatusID(ByRef f_IOStatusID)
    If Not IsNumeric(f_IOStatusID) Then
      ' ######
      ' Print an error and die
    End If

    ' ######
    ' Validate in tblLists

    m_IOStatusID = f_IOStatusID
    Set m_IOStatus = Nothing
  End Property

  Public Property Let IOStatusDate(ByRef f_IOStatusDate)
    If Not IsDate(f_IOStatusDate) Then
      ' #######
      ' Print an error and die
    End If
    m_IOStatusDate = f_IOStatusDate
  End Property

  Public Property Get IOStatusDate() ' As Date
    IOStatusDate = m_IOStatusDate
  End Property

  Public Property Get IOStatusText() ' As String
    IOStatusText = m_IOStatusText
  End Property

  Public Property Let IOStatusText(ByRef f_IOStatusText)
    m_IOStatusText = Left(Trim(f_IOStatusText), 255)
  End Property

  Public Property Get TZOffSet() ' As Int
    TZOffSet = m_TZOffSet
  End Property

  Public Property Let TZOffSet(ByRef f_TZOffSet)
    If Not IsNumeric(f_TZOffSet) Then
      ' ######
      ' Print an error and die
    End If

    m_TZOffSet = f_TZOffSet
  End Property

  Public Property Let UserPermMask(ByRef f_UserPermMask)
    m_UserPermMask = f_UserPermMask OR &H0000000000000000
  End Property

  Public Property Get UserPermMask()  ' As Binary
    UserPermMask = m_UserPermMask
  End Property

  Public Property Get Role() ' As clsRole
    If Not IsObject(m_Role) Then
      Dim oRole, blnTemp
      Set oRole = New clsRole
      If IsNumeric(m_RoleID) Then
        oRole.ID = m_RoleID
        blnTemp = oRole.Load
      End If
      Set m_Role = oRole
    End If
    Set Role = m_Role
  End Property

  Public Property Let RoleID(ByRef f_RoleID)
    Dim oRole
    Set oRole = New clsRole
    oRole.ID = f_RoleID
    If Not oRole.Load Then
      ' ######
      ' Raise Error
    End If
    m_RoleID = f_RoleID
    Set m_Role = oRole
  End Property

  Public Property Get RoleID() ' As Long
    RoleID = m_RoleID
  End Property

  Public Property Let OfficePhone(ByRef f_OfficePhone)
    m_OfficePhone = Left(Trim(f_OfficePhone), 20)
  End Property

  Public Property Get OfficePhone() ' As String
    OfficePhone = m_OfficePhone
  End Property

  Public Property Let HomePhone(ByRef f_HomePhone)
    m_HomePhone = Left(Trim(f_HomePhone), 20)
  End Property

  Public Property Get HomePhone() ' As String
    HomePhone = m_HomePhone
  End Property

  Public Property Let MobilePhone(ByRef f_MobilePhone)
    m_MobilePhone = Left(Trim(f_MobilePhone), 20)
  End Property

  Public Property Get MobilePhone() ' As String
    MobilePhone = m_MobilePhone
  End Property

  Public Property Let JobTitle(ByRef f_JobTitle)
    m_JobTitle = Left(Trim(f_JobTitle), 80)
  End Property

  Public Property Get JobTitle() ' As String
    JobTitle = m_JobTitle
  End Property

  Public Property Get JobFunction() ' As clsListItem
    If IsObject(m_JobFunction) Then
      Set JobFunction = m_JobFunction
    Else
      ' ######
      ' Create and return JobFunction
    End If
  End Property

  Public Property Get JobFunctionID() ' As clsListItem
    JobFunctionID = m_JobFunctionID
  End Property

  Public Property Let JobFunctionID(ByRef f_JobFunctionID)
    If Not IsNumeric(f_JobFunctionID) Then
      ' ######
      ' Print an error message and die
    End If

    ' ######
    ' Verify against tblLists

    m_JobFunctionID = f_JobFunctionID
    Set m_JobFunction = Nothing
  End Property

  Public Property Let sResume(ByRef f_Resume)
    m_Resume = Trim(f_Resume)
  End Property

  Public Property Get sResume() ' As String
    sResume = m_Resume
  End Property

  Public Property Let Email(ByRef f_Email)
    m_Email = Left(Trim(f_Email), 40)
  End Property

  Public Property Get Email() ' As String
    Email = m_Email
  End Property

  Public Property Let PagerEmail(ByRef f_PagerEmail)
    m_PagerEmail = Left(Trim(f_PagerEmail), 40)
  End Property

  Public Property Get PagerEmail() ' As String
    PagerEmail = m_PagerEmail
  End Property

  Public Property Let OfficeLocation(ByRef f_OfficeLocation)
    m_OfficeLocation = Left(Trim(f_OfficeLocation), 80)
  End Property

  Public Property Get OfficeLocation() ' As String
    OfficeLocation = m_OfficeLocation
  End Property

  Public Property Let Notes(ByRef f_Notes)
    m_Notes = Trim(f_Notes)
  End Property

  Public Property Get Notes() ' As String
    Notes = m_Notes
  End Property

  Public Property Get PhotoFile() ' As clsFile
    If IsObject(m_PhotoFile) Then
      Set PhotoFile = m_PhotoFile
    Else
      ' ######
      ' Get File object
    End If
  End Property

  Public Property Get PhotoFileID() ' As Long
    PhotoFileID = m_PhotoFileID
  End Property

  Public Property Let PhotoFileID(ByRef f_PhotoFileID)
    If Not IsNumeric(f_PhotoFileID) Then
      ' ######
      ' Print error and die
    End If

    ' ######
    ' Validate from tblFiles

    m_PhotoFileID = f_PhotoFileID
  End Property

  Public Property Get Created() ' As Date
    Created = m_Created
  End Property

  Public Property Let LastAccess(ByRef f_LastAccess)
    If Not IsDate(f_LastAccess) Then
      ' #######
      ' Print an error and die
    End If
    m_LastAccess = f_LastAccess
  End Property

  Public Property Get LastAccess() ' As Date
    LastAccess = m_LastAccess
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

  ' #################
  ' Public Methods
  ' #################

  Public Function Load() ' As Boolean
    ' One of the following is required to load:
    ' SourceRS
    ' ID
    ' UserName

    Dim blnLoad, rsContact, strQuery, blnUseSourceRS
    blnLoad = True
    blnUseSourceRS = False
    If IsObject(SourceRS) Then
      If SourceRS.State <> adStateClosed Then
        Set rsContact = SourceRS
        blnUseSourceRS = True
      Else
        blnLoad = False
        m_LastError = "SourceRS is closed."
      End If
    Else
      If Len(m_UserName) > 0 Then
        strQuery = "SELECT * from tblContacts WHERE UserName = '" & m_UserName & "'"
'      End If
	  ElseIf IsNumeric(m_ID) And Not IsEmpty(m_ID) Then
        strQuery = "SELECT * from tblContacts WHERE ContactPK = " & m_ID
      End If
      
      If Len(strQuery) > 0 Then
        Set rsContact = Server.CreateObject("ADODB.RecordSet")
        rsContact.Open strQuery, m_cnnDB
      Else
        blnLoad = False
        m_LastError = "Missing ID or username."
      End If
      
    End If
    If blnLoad Then
      If Not rsContact.EOF And Not rsContact.BOF Then
        On Error Resume Next
        m_ID = rsContact("ContactPK")
        m_OrgID = rsContact("OrgFK")
        m_FName = rsContact("FName")
        m_LName = rsContact("LName")
        m_DeptID = rsContact("DeptFK")
        m_ContactTypeID = rsContact("ContactTypeFK")
        m_LangID = rsContact("LangFK")
        m_IsActive = rsContact("IsActive")
        m_UserName = rsContact("UserName")
        m_Password = rsContact("PW")
        m_IOStatusID = rsContact("IOStatusFK")
        m_IOStatusDate = rsContact("IOStatusDate")
        m_IOStatusText = rsContact("IOStatusText")
        m_TZOffset = rsContact("TZOffset")
        m_UserPermMask = rsContact("UserPermMask")
        m_RoleID = rsContact("RoleFK")
        m_OfficePhone = rsContact("OfficePhone")
        m_HomePhone = rsContact("HomePhone")
        m_MobilePhone = rsContact("MobilePhone")
        m_JobTitle = rsContact("JobTitle")
        m_JobFunctionID = rsContact("JobFunction")
        m_Resume = rsContact("Resume")
        m_EMail = rsContact("EMail")
        m_OfficeLocation = rsContact("OfficeLocation")
        m_Notes = rsContact("Notes")
        m_Created = rsContact("Created")
        m_LastAccess = rsContact("LastAccess")
        m_PagerEmail = rsContact("PagerEMail")
        m_PhotoFileID = rsContact("PhotoFileFK")
        m_LastUpdate = rsContact("LastUpdate")
        m_LastUpdateBy = rsContact("LastUpdateByFK")
        On Error Goto 0
        If Err.Number <> 0 Then
          blnLoad = False
          m_LastError = "Error retrieving data -- " & Err.Source & ": " & Err.Description
        End If
      Else
        blnLoad = False
        m_LastError = "No matching contacts found."
      End If
    End If
    
    If Not blnUseSourceRS And blnLoad Then
      rsContact.Close
      Set rsContact = Nothing
    End If
    
    If blnLoad Then
      m_IsValid = True
    End If
    
    Load = blnLoad
  End Function

  Public Function Update() ' As Boolean
    Dim blnUpdate, rsContact
    blnUpdate = True
    Set rsContact = Server.CreateObject("ADODB.RecordSet")
    If m_IsValid Then
      rsContact.Open "SELECT * FROM tblContacts WHERE ContactPK = " & m_ID, m_cnnDB, _
        adOpenKeyset, adLockOptimistic, adCmdText
    Else
      rsContact.Open "tblContacts", m_cnnDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect
      rsContact.AddNew
      
      rsContact.Fields("Created") = m_Created
      
    End If
    With rsContact
      .Fields("OrgFK") = m_OrgID
      .Fields("FName") = m_FName
      .Fields("LName") = m_LName
      .Fields("DeptFK") = m_DeptID
      .Fields("ContactTypeFK") = m_ContactTypeID
      .Fields("LangFK") = m_LangID
      .Fields("IsActive") = m_IsActive
      .Fields("UserName") = m_UserName
      .Fields("PW") = m_Password
      .Fields("IOStatusFK") = m_IOStatusID
      .Fields("IOStatusDate") = m_IOStatusDate
      .Fields("IOStatusText") = m_IOStatusText
      .Fields("TZOffset") = m_TZOffset
      .Fields("UserPermMask") = m_UserPermMask
      .Fields("RoleFK") = m_RoleID
      .Fields("OfficePhone") = m_OfficePhone
      .Fields("HomePhone") = m_HomePhone
      .Fields("MobilePhone") = m_MobilePhone
      .Fields("JobTitle") = m_JobTitle
      .Fields("JobFunction") = m_JobFunctionID
      .Fields("Resume") = m_Resume
      .Fields("EMail") = m_EMail
      .Fields("OfficeLocation") = m_OfficeLocation
      .Fields("Notes") = m_Notes
'      .Fields("LastAccess") = m_LastAccess
      .Fields("PagerEMail") = m_PagerEMail
      .Fields("PhotoFileFK") = m_PhotoFileID
      .Fields("LastUpdate") = m_LastUpdate
      .Fields("LastUpdateByFK") = m_LastUpdateByID
      .Update
      m_ID = .Fields("ContactPK")
    End With
    rsContact.Close
    Set rsContact = Nothing
    m_IsValid = True
    Update = blnUpdate
  End Function

  Public Function Delete() ' As Boolean
    Dim blnDelete
    blnDelete = True
    m_IsActive = False
    If m_IsValid Then
      m_cnnDB.Execute "UPDATE tblContacts SET IsActive = " & m_IsActive & _
        " WHERE ContactPK = " & m_ID
    End If
    Delete = blnDelete
  End Function

  Public Function CheckPassword(ByVal f_Password) ' As Boolean
    ' #######
    ' Encrypt password in MD5
    If f_Password = m_Password Then
      CheckPassword = True
    Else
      CheckPassword = False
    End If
  End Function

  Public Sub LogLastAccess() ' As Boolean


	m_cnnDB.Execute "UPDATE tblContacts " &_
					"SET LastAccess='" & CStr(Right("0" & Month(Now()), 2) & "/" & Right("0" & Day(Now()), 2) & "/" & Year(Now()) & " " & FormatDateTime(Now(), vbShortTime)) & "' " &_
					"WHERE ContactPK=" & m_ID

  End Sub

  ' ##################
  ' Private Methods
  ' ##################
  Private Sub Class_Initialize()
  
    If IsObject(cnnDB) Then
      Set m_cnnDB = cnnDB
    End If
    m_IsValid = False

'    m_FName = "Unknown"
'    m_LName = "Unknown"
'    m_LangID = 1  ' English
'    m_UserName = "Unknown"
'    m_IsUser = True
    m_TZOffset = 0

    ' Set Default Contact Properties

    m_IsActive = True
    m_UserPermMask = &H0000000000000000
    m_RoleID = Application("DEFAULT_ROLE")
    m_Created = CStr(Right("0" & Month(Now()), 2) & "/" & Right("0" & Day(Now()), 2) & "/" & Year(Now()) & " " & FormatDateTime(Now(), vbShortTime))

  End Sub

  Private Sub Class_Terminate()
  End Sub

End Class
%>
