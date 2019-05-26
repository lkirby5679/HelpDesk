<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: clsCase.asp
'  Date:     $Date: 2004/03/11 06:29:08 $
'  Version:  $Revision: 1.2 $
'  Purpose:  Class object to assist when working with Cases
' ----------------------------------------------------------------------------------

Class clsCase
  ' ##################
  ' Private Properties
  ' ##################
  Private m_ID
  Private m_ContactID, m_Contact, m_RepID, m_Rep, m_GroupID, m_Group
  Private m_StatusID, m_Status, m_CatID, m_Cat, m_PriorityID, m_Priority
  Private m_CaseTypeID, m_CaseType, m_Title, m_Description, m_Resolution
  Private m_AltEMail, m_RaisedDate, m_ClosedDate, m_IsActive
  Private m_EnteredByID, m_EnteredBy, m_CaseNotes, m_CaseFiles
  Private m_LastUpdate, m_LastUpdateBy, m_LastUpdateByID
  Private m_IsValid, m_LastError, m_cnnDB, m_CC, m_Dept, m_DeptID

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

  Public Property Get Contact() ' As clsContact
    If Not IsObject(m_Contact) Then
      Dim oContact, blnTemp
      Set oContact = New clsContact
      If IsNumeric(m_ContactID) Then
        oContact.ID = m_ContactID
        blnTemp = oContact.Load
      End If
      Set m_Contact = oContact
    End If
    Set Contact = m_Contact
  End Property

  Public Property Get ContactID() ' As Long
    ContactID = m_ContactID
  End Property

  Public Property Let ContactID(ByRef f_ContactID)

    Dim objContact
    
    If IsNumeric(f_ContactID) And Not IsEmpty(f_ContactID) And f_ContactID > 0 Then
    
      Set objContact = New clsContact

      objContact.ID = f_ContactID

      If Not objContact.Load Then
        ' Item not found, Raise Error
      Else
        m_ContactID = f_ContactID
        Set m_Contact = objContact
      End If
      
      Set objContact = Nothing

    Else
    
      ' Invalid ID, Raise Error
    
    End If
    
  End Property

  Public Property Get Rep() ' As clsContact
    If Not IsObject(m_Rep) Then
      Dim oRep, blnTemp
      Set oRep = New clsContact
      If IsNumeric(m_RepID) Then
        oRep.ID = m_RepID
        blnTemp = oRep.Load
      End If
      Set m_Rep = oRep
    End If
    Set Rep = m_Rep
  End Property

  Public Property Get RepID() ' As Long
    RepID = m_RepID
  End Property

  Public Property Let RepID(ByRef f_RepID)

    Dim objRep
    
    If IsNumeric(f_RepID) And Not IsEmpty(f_RepID) And f_RepID > 0 Then
    
      Set objRep = New clsContact

      objRep.ID = f_RepID

      If Not objRep.Load Then
        ' Item not found, Raise Error
      Else
        m_RepID = f_RepID
        Set m_Rep = objRep
      End If
      
      Set objRep = Nothing
      
    Else
    
      ' This step caters for the time we set the Assignment to unassigned
      If f_RepID = 0 Then
        m_RepID = Empty
      Else
        ' Invalid ID, Raise Error
      End If
    
    End If

  End Property

  Public Property Get Group() ' As clsGroup
  
	Dim objGroup
  
    If Not IsObject(m_Group) Then
    
      If IsNumeric(m_GroupID) Then
      
        Set objGroup = New clsGroup
        
        objGroup.ID = m_GroupID
        
        If Not objGroup.Load Then
          ' Raise error
        Else
          Set m_Group = objGroup
        End If
        
        Set objGroup = Nothing
        
      Else
      
        ' Invalid Group ID, Raise error
      
      End If
      
    Else
    
	  ' m_Group already loaded
      
    End If
    
    Set Group = m_Group
    
  End Property

  Public Property Get GroupID() ' As Long
    GroupID = m_GroupID
  End Property

  Public Property Let GroupID(ByRef f_GroupID)

    Dim objGroup
    
    If IsNumeric(f_GroupID) And Not IsEmpty(f_GroupID) And f_GroupID > 0 Then
    
      Set objGroup = New clsGroup

      objGroup.ID = f_GroupID

      If Not objGroup.Load Then
        ' Item not found, Raise Error
      Else
         m_GroupID = f_GroupID
         Set m_Group = objGroup
      End If
      
      Set objGroup = Nothing

    Else
    
      ' Invalid ID, Raise Error
    
    End If
    
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

  Public Property Get Status() ' As clsListItem
    If Not IsObject(m_Status) Then
      Dim oStatus, blnTemp
      Set oStatus = New clsListItem
      If IsNumeric(m_StatusID) Then
        oStatus.ID = m_StatusID
        blnTemp = oStatus.Load
      End If
      Set m_Status = oStatus
    End If
    Set Status = m_Status
  End Property

  Public Property Get StatusID() ' As Long
    StatusID = m_StatusID
  End Property

  Public Property Let StatusID(ByRef f_StatusID)

    Dim objStatus
    
    If IsNumeric(f_StatusID) And Not IsEmpty(f_StatusID) And f_StatusID > 0 Then
    
      Set objStatus = New clsListItem

      objStatus.ID = f_StatusID

      If Not objStatus.Load Then
        ' Item not found, Raise Error
      Else
         m_StatusID = f_StatusID
         Set m_Status = objStatus
      End If
      
      Set objStatus = Nothing
      
    Else
    
      ' Invalid ID, Raise Error
    
    End If

  End Property

  Public Property Get Cat() ' As clsCategory
    If Not IsObject(m_Cat) Then
      Dim oCat, blnTemp
      Set oCat = New clsCategory
      If IsNumeric(m_CatID) Then
        oCat.ID = m_CatID
        blnTemp = oCat.Load
      End If
      Set m_Cat = oCat
    End If
    Set Cat = m_Cat
  End Property

  Public Property Get CatID() ' As Long
    CatID = m_CatID
  End Property

  Public Property Let CatID(ByRef f_CatID)

    Dim objCat
    
    If IsNumeric(f_CatID) And Not IsEmpty(f_CatID) And f_CatID > 0 Then
    
      Set objCat = New clsCategory

      objCat.ID = f_CatID

      If Not objCat.Load Then
        ' Item not found, Raise Error
      Else
        m_CatID = f_CatID
        Set m_Cat = objCat
      End If
      
      Set objCat = Nothing
      
    Else
    
      ' Invalid ID, Raise Error
    
    End If

  End Property

  Public Property Get Priority() ' As clsListItem
    If Not IsObject(m_Priority) Then
      Dim oPriority, blnTemp
      Set oPriority = New clsListItem
      If IsNumeric(m_PriorityID) Then
        oPriority.ID = m_PriorityID
        blnTemp = oPriority.Load
      End If
      Set m_Priority = oPriority
    End If
    Set Priority = m_Priority
  End Property

  Public Property Get PriorityID() ' As Long
    PriorityID = m_PriorityID
  End Property

  Public Property Let PriorityID(ByRef f_PriorityID)

    Dim objPriority
    
    If IsNumeric(f_PriorityID) And Not IsEmpty(f_PriorityID) And f_PriorityID > 0 Then
    
      Set objPriority = New clsListItem

      objPriority.ID = f_PriorityID

      If Not objPriority.Load Then
        ' Item not found, Raise Error
      Else
         m_PriorityID = f_PriorityID
         Set m_Priority = objPriority
      End If
      
      Set objPriority = Nothing

    Else
    
      ' Invalid ID, Raise Error
    
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

    Dim objCaseType
    
    If IsNumeric(f_CaseTypeID) And Not IsEmpty(f_CaseTypeID) And f_CaseTypeID > 0 Then
    
      Set objCaseType = New clsCaseType

      objCaseType.ID = f_CaseTypeID

      If Not objCaseType.Load Then
        ' Item not found, Raise Error
      Else
         m_CaseTypeID = f_CaseTypeID
         Set m_CaseType = objCaseType
      End If
      
      Set objCaseType = Nothing

    Else
    
      ' Invalid ID, Raise Error
    
    End If
    
  End Property

  Public Property Get Title() ' As String
    Title = m_Title
  End Property

  Public Property Let Title(ByRef f_Title)
    m_Title = Left(Trim(f_Title), 80)
  End Property

  Public Property Get Cc() ' As String
    CC = m_Cc
  End Property

  Public Property Let Cc(ByRef f_Cc)
    m_Cc = Left(Trim(f_CC), 36)
  End Property

  Public Property Get Description() ' As String
    Description = m_Description
  End Property

  Public Property Let Description(ByRef f_Description)
    m_Description = Trim(f_Description)
  End Property

  Public Property Get Resolution() ' As String
    Resolution = m_Resolution
  End Property

  Public Property Let Resolution(ByRef f_Resolution)
    m_Resolution = Trim(f_Resolution)
  End Property

  Public Property Get AltEMail() ' As String
    AltEMail = m_AltEMail
  End Property

  Public Property Let AltEMail(ByRef f_AltEMail)
    m_AltEMail = Left(Trim(f_AltEMail), 40)
  End Property

  Public Property Get RaisedDate() ' As Date
    RaisedDate = m_RaisedDate
  End Property

  Public Property Let RaisedDate(ByRef f_RaisedDate)
    If IsDate(f_RaisedDate) Then
      m_RaisedDate = f_RaisedDate
    End If
  End Property

  Public Property Get ClosedDate() ' As Date
    ClosedDate = m_ClosedDate
  End Property

  Public Property Let ClosedDate(ByRef f_ClosedDate)
    If IsDate(f_ClosedDate) Then
      m_ClosedDate = f_ClosedDate
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

  Public Property Get EnteredBy() ' As clsContact
    If Not IsObject(m_EnteredBy) Then
      Dim oEnteredBy, blnTemp
      Set oEnteredBy = New clsContact
      If IsNumeric(m_EnteredByID) Then
        oEnteredBy.ID = m_EnteredByID
        blnTemp = oEnteredBy.Load
      End If
      Set m_EnteredBy = oEnteredBy
    End If
    Set EnteredBy = m_EnteredBy
  End Property

  Public Property Get CaseNotes() ' As clsCollection of clsNote
    If m_IsValid Then
      If Not IsObject(m_CaseNotes) Then
        Dim f_CN, blnTemp
        Set f_CN = New clsCollection
        f_CN.CollectionType = f_CN.clNote
        f_CN.Query = "SELECT * FROM tblNotes WHERE CaseFK = " & m_ID & " ORDER BY NotePK ASC"
'        f_CN.Where = "CaseFK = " & m_ID
'        f_CN.OrderBy = "NotePK ASC"
        blnTemp = f_CN.Load
        Set m_CaseNotes = f_CN
      End If
      Set CaseNotes = m_CaseNotes
    Else
      Set CaseNotes = Nothing
    End If
  End Property

  Public Property Get EnteredByID() ' As Long
    EnteredByID = m_EnteredByID
  End Property

  Public Property Let EnteredByID(ByRef f_EnteredByID)

    Dim objEnteredBy
    
    If IsNumeric(f_EnteredByID) And Not IsEmpty(f_EnteredByID) And f_EnteredByID > 0 Then
    
      Set objEnteredBy = New clsContact

      objEnteredBy.ID = f_EnteredByID

      If Not objEnteredBy.Load Then
        ' Item not found, Raise Error
      Else
         m_EnteredByID = f_EnteredByID
         Set m_EnteredBy = objEnteredBy
      End If
      
      Set objEnteredBy = Nothing

    Else
    
      ' Invalid ID, Raise Error
    
    End If
    
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

    Dim objLastUpdateBy
    
    If IsNumeric(f_LastUpdateByID) And Not IsEmpty(f_LastUpdateByID) And f_LastUpdateByID > 0 Then
    
      Set objLastUpdateBy = New clsContact

      objLastUpdateBy.ID = f_LastUpdateByID

      If Not objLastUpdateBy.Load Then
        ' Item not found, Raise Error
      Else
        m_LastUpdateByID = f_LastUpdateByID
        Set m_LastUpdateBy = objLastUpdateBy
      End If
      
      Set objLastUpdateBy = Nothing

    Else
    
      ' Invalid ID, Raise Error
    
    End If
    
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

    Dim blnLoad, rsCase, strQuery, blnUseSourceRS
    blnLoad = True
    blnUseSourceRS = False
    If IsObject(SourceRS) Then
      If SourceRS.State <> adStateClosed Then
        Set rsCase = SourceRS
        blnUseSourceRS = True
      Else
        blnLoad = False
        m_LastError = "SourceRS is closed."
      End If
    Else
      If IsNumeric(m_ID) And Not IsEmpty(m_ID) Then
        strQuery = "SELECT * FROM tblCases WHERE CasePK = " & m_ID
        Set rsCase = Server.CreateObject("ADODB.RecordSet")
        rsCase.Open strQuery, m_cnnDB
      Else
        blnLoad = False
        m_LastError = "Missing ID."
      End If
    End If
    If blnLoad Then
      If Not rsCase.EOF And Not rsCase.BOF Then
        On Error Resume Next
        m_ID = rsCase("CasePK")
        m_ContactID = rsCase("ContactFK")
        m_RepID = rsCase("RepFK")
        m_GroupID = rsCase("GroupFK")
        m_StatusID = rsCase("StatusFK")
        m_CatID = rsCase("CatFK")
        m_PriorityID = rsCase("PriorityFK")
        m_CaseTypeID = rsCase("CaseTypeFK")
        m_Title = rsCase("Title")
        m_Description = rsCase("Description")
        m_Resolution = rsCase("Resolution")
        m_AltEMail = rsCase("AltEMail")
        m_Cc = rsCase("Cc")
        m_DeptID = rsCase("DeptFK")
        m_RaisedDate = rsCase("RaisedDate")
        m_ClosedDate = rsCase("ClosedDate")
        m_IsActive = rsCase("IsActive")
        m_EnteredByID = rsCase("EnteredByFK")
        m_LastUpdate = rsCase("LastUpdate")
        m_LastUpdateByID = rsCase("LastUpdateByFK")
        Dim blnTemp
        ' Pre-load Contact
        'Response.Write "Preload Contact. --" & rsCase("Contact.UserName")
        'Set m_Contact = New clsContact
        'Set m_Contact.SourceRS = rsCase
        'm_Contact.SourceRSTable = "Contact"
        'blnTemp = m_Contact.Load
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
    If Not blnUseSourceRS And blnLoad Then
      rsCase.Close
      Set rsCase = Nothing
    End If
    If blnLoad Then
      m_IsValid = True
    End If
    Load = blnLoad
  End Function

  Public Function Update() ' As Boolean
    Dim blnUpdate, rsCase
    blnUpdate = True
    Set rsCase = Server.CreateObject("ADODB.RecordSet")
    If m_IsValid Then
      rsCase.Open "SELECT * FROM tblCases WHERE CasePK = " & m_ID, m_cnnDB, _
        adOpenKeyset, adLockOptimistic, adCmdText
    Else
      rsCase.Open "tblCases", m_cnnDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect
      rsCase.AddNew
    End If
    With rsCase
      .Fields("ContactFK") = m_ContactID
      .Fields("RepFK") = m_RepID
      .Fields("GroupFK") = m_GroupID
      .Fields("StatusFK") = m_StatusID
      .Fields("CatFK") = m_CatID
      .Fields("PriorityFK") = m_PriorityID
      .Fields("CaseTypeFK") = m_CaseTypeID
      .Fields("Title") = m_Title
      .Fields("Description") = m_Description
      .Fields("Resolution") = m_Resolution
      .Fields("AltEMail") = m_AltEMail
      .Fields("Cc") = m_Cc
      .Fields("DeptFK") = m_DeptID
      .Fields("RaisedDate") = m_RaisedDate
      .Fields("ClosedDate") = m_ClosedDate
      .Fields("IsActive") = m_IsActive
      .Fields("EnteredByFK") = m_EnteredByID
      .Fields("LastUpdate") = m_LastUpdate
      .Fields("LastUpdateByFK") = m_LastUpdateByID
      .Update
      m_ID = .Fields("CasePK")
    End With
    rsCase.Close
    Set rsCase = Nothing
    m_IsValid = True
    Update = blnUpdate
  End Function

  Public Function Delete() ' As Boolean
    Dim blnDelete
    blnDelete = True
    m_IsActive = False
    If m_IsValid Then
      m_cnnDB.Execute "UPDATE tblCases SET IsActive = " & m_IsActive & _
        " WHERE CasePK = " & m_ID, , adExecuteNoRecords
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

    ' Set Default Case Properties
    
'    m_IsActive = True
'    m_RaisedDate = Now
'    m_LastUpdate = Now
    
  End Sub

  Private Sub Class_Terminate()
  End Sub

End Class
%>