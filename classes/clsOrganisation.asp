<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: clsOrganisation.asp
'  Date:     $Date: 2004/03/11 06:29:08 $
'  Version:  $Revision: 1.2 $
'  Purpose:  Class object to assist when working with Organisations
' ----------------------------------------------------------------------------------

Class clsOrganisation
  ' ##################
  ' Private Properties
  ' ##################
  Private m_ID, m_OrgType, m_OrgTypeID, m_OrgShortName, m_OrgName
  Private m_IsActive, m_Password, m_PrimaryContact, m_PrimaryContactID
  Private m_OfficeLocation, m_Phone, m_Fax, m_Email, m_MailAddress
  Private m_CourierAddress, m_City, m_State, m_Country, m_Notes
  Private m_LastUpdate, m_LastUpdateBy, m_LastUpdateByID

  Private m_IsValid, m_LastError, m_cnnDB

  ' #################
  ' Public Properties
  ' #################
  Public SourceRS  ' As RecordSet

  Public Property Get ID()  ' As Long
    ID = m_ID
  End Property

  Public Property Let ID(ByRef f_ID)
    If IsNumeric(f_ID) Then
      m_ID = f_ID
    End If
  End Property

  Public Property Get OrgType() ' As clsListItem

    Dim oOrgType, blnTemp

    If Not IsObject(m_OrgType) Then

      Set oOrgType = New clsListItem
      If IsNumeric(m_OrgTypeID) Then
        oOrgType.ID = m_OrgTypeID
        blnTemp = oOrgType.Load
      End If
      Set m_OrgType = oOrgType

    End If
    Set OrgType = m_OrgType

  End Property

  Public Property Get OrgTypeID() ' As Long
    OrgTypeID = m_OrgTypeID
  End Property

  Public Property Let OrgTypeID(ByRef f_OrgTypeID)

    Dim oOrgType
    
    If Not IsNumeric(f_OrgTypeID) Then
      ' ######
      ' Print error and die
    Else
      ' ######
      ' Validate from tblLists
      Set oOrgType = New clsListItem
      
      oOrgType.ID = f_OrgTypeID
      
      If Not oOrgType.Load Then
      
         ' Raise Error
		
	  Else

         m_OrgTypeID = f_OrgTypeID
         Set m_OrgType = oOrgType
	  
	  End If
	  
'	  Set oOrgType = Nothing

    End If

  End Property

  Public Property Get OrgShortName()  ' As String
    OrgShortName = m_OrgShortName
  End Property

  Public Property Let OrgShortName(ByRef f_OrgShortName)
    m_OrgShortName = Left(Trim(f_OrgShortName), 16)
  End Property

  Public Property Get OrgName() ' As String
    OrgName = m_OrgName
  End Property

  Public Property Let OrgName(ByRef f_OrgName)
    m_OrgName = Left(Trim(f_OrgName), 80)
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

  Public Property Get Password()  ' As Boolean
    Password = m_Password
  End Property

  Public Property Let Password(ByVal f_password)
    ' ######
    ' MD5 encrypt

    m_Password = f_password
  End Property


  Public Property Get PrimaryContact()  ' As clsContact
    
    Dim oPrimaryContact

    Set oPrimaryContact = New clsContact

    If IsNumeric(m_PrimaryContactID) Then
      
        oPrimaryContact.ID = m_PrimaryContactID
        
      If Not oPrimaryContact.Load Then
	    ' Not loaded
      Else
	    Set PrimaryContact = oPrimaryContact
      End If
        
        
    Else
      
       ' Contact isn't in database
         
  	  Set PrimaryContact = oPrimaryContact
        
    End If


    Set oPrimaryContact = Nothing

  End Property


  Public Property Get PrimaryContactID()  'As Long
    PrimaryContactID = m_PrimaryContactID
  End Property

  Public Property Let PrimaryContactID(ByRef f_PrimaryContactID)
    Dim oContact
    
    Set oContact = New clsContact
    oContact.ID = f_PrimaryContactID
    
    If Not oContact.Load Then
    
      ' Special case for a Primary Contact not selected.
      If f_PrimaryContactID = 0 Then
         m_PrimaryContactID = Empty
      Else
         ' Raise Error
      End If
    
    Else
    
       m_PrimaryContactID = f_PrimaryContactID
       Set m_PrimaryContact = oContact
    
    End If
    
    Set oContact = Nothing
  End Property

  Public Property Get OfficeLocation()  ' As String
    OfficeLocation = m_OfficeLocation
  End Property

  Public Property Let OfficeLocation(ByRef f_OfficeLocation)
    m_OfficeLocation = Left(Trim(f_OfficeLocation), 24)
  End Property

  Public Property Get Phone() ' As String
    Phone = m_Phone
  End Property

  Public Property Let Phone(ByRef f_Phone)
    m_Phone = Left(Trim(f_Phone), 20)
  End Property

  Public Property Get Fax() ' As String
    Fax = m_Fax
  End Property

  Public Property Let Fax(ByRef f_Fax)
    m_Fax = Left(Trim(f_Fax), 20)
  End Property

  Public Property Get Email() ' As String
    Email = m_Email
  End Property

  Public Property Let Email(ByRef f_Email)
    m_Email = Left(Trim(f_Email), 32)
  End Property

  Public Property Get MailAddress() ' As String
    MailAddress = m_MailAddress
  End Property

  Public Property Let MailAddress(ByRef f_MailAddress)
    m_MailAddress = Left(Trim(f_MailAddress), 160)
  End Property

  Public Property Get CourierAddress() ' As String
    CourierAddress = m_CourierAddress
  End Property

  Public Property Let CourierAddress(ByRef f_CourierAddress)
    m_CourierAddress = Left(Trim(f_CourierAddress), 160)
  End Property
  
  Public Property Get City() ' As String
    City = m_City
  End Property

  Public Property Let City(ByRef f_City)
    m_City = Left(Trim(f_City), 40)
  End Property

  Public Property Get State() ' As String
    State = m_State
  End Property

  Public Property Let State(ByRef f_State)
    m_State = Left(Trim(f_State), 40)
  End Property

  Public Property Get Country() ' As String
    Country = m_Country
  End Property

  Public Property Let Country(ByRef f_Country)
    m_Country = Left(Trim(f_Country), 40)
  End Property

  Public Property Get Notes() ' As String
    Notes = m_Notes
  End Property

  Public Property Let Notes(ByRef f_Notes)
    m_Notes = Trim(f_Notes)
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
  Public Function Load() ' As Boolean
    ' One of the following is required to load:
    ' SourceRS
    ' ID

    Dim blnLoad, rsOrg, strQuery, blnUseSourceRS
    blnLoad = True
    blnUseSourceRS = False
    If IsObject(SourceRS) Then
      If SourceRS.State <> adStateClosed Then
        Set rsOrg = SourceRS
        blnUseSourceRS = True
      Else
        blnLoad = False
        m_LastError = "SourceRS is closed."
      End If
    Else
      If IsNumeric(m_ID) Then
        strQuery = "SELECT * FROM tblOrganisations WHERE OrgPK = " & m_ID
        Set rsOrg = Server.CreateObject("ADODB.RecordSet")
        rsOrg.Open strQuery, m_cnnDB
      Else
        blnLoad = False
        m_LastError = "Missing ID."
      End If
    End If
    If blnLoad Then
      If Not rsOrg.EOF And Not rsOrg.BOF Then
        On Error Resume Next
        m_ID = rsOrg("OrgPK")
        m_OrgTypeID = rsOrg("OrgTypeFK")
        m_OrgShortName = rsOrg("OrgShortName")
        m_OrgName = rsOrg("OrgName")
        m_Password = rsOrg("PW")
        m_PrimaryContactID = rsOrg("PrimaryContactFK")
        m_OfficeLocation = rsOrg("OfficeLocation")
        m_Phone = rsOrg("Phone")
        m_Fax = rsOrg("Fax")
        m_Email = rsOrg("Email")
        m_MailAddress = rsOrg("MailAddress")
        m_CourierAddress = rsOrg("CourierAddress")
        m_City = rsOrg("City")
        m_State = rsOrg("State")
        m_Country = rsOrg("Country")
        m_Notes = rsOrg("Notes")
        m_IsActive = rsOrg("IsActive")
        m_LastUpdate = rsOrg("LastUpdate")
        m_LastUpdateBy = rsOrg("LastUpdateByFK")
        On Error Goto 0
        If Err.Number <> 0 Then
          blnLoad = False
          m_LastError = "Error retrieving data -- " & Err.Source & ": " & Err.Description
        End If
      Else
        blnLoad = False
        m_LastError = "No matching organizations found."
      End If
    End If
    If Not blnUseSourceRS Then
      rsOrg.Close
      Set rsOrg = Nothing
    End If
    If blnLoad Then
      m_IsValid = True
    End If
    Load = blnLoad
  End Function

  Public Function Update() ' As Boolean
    Dim blnUpdate, rsOrg
    blnUpdate = True
    Set rsOrg = Server.CreateObject("ADODB.RecordSet")
    If m_IsValid Then
      rsOrg.Open "SELECT * FROM tblOrganisations WHERE OrgPK = " & m_ID, m_cnnDB, _
        adOpenKeyset, adLockOptimistic, adCmdText
    Else
      rsOrg.Open "tblOrganisations", m_cnnDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect
      rsOrg.AddNew
    End If
    With rsOrg
      .Fields("OrgTypeFK") = m_OrgTypeID
      .Fields("OrgShortName") = m_OrgShortName
      .Fields("OrgName") = m_OrgName
      .Fields("PW") = m_Password
      .Fields("PrimaryContactFK") = m_PrimaryContactID
      .Fields("OfficeLocation") = m_OfficeLocation
      .Fields("Phone") = m_Phone
      .Fields("Fax") = m_Fax
      .Fields("Email") = m_Email
      .Fields("MailAddress") = m_MailAddress
      .Fields("CourierAddress") = m_CourierAddress
      .Fields("City") = m_City
      .Fields("State") = m_State
      .Fields("Country") = m_Country
      .Fields("Notes") = m_Notes
      .Fields("IsActive") = m_IsActive
      .Fields("LastUpdate") = m_LastUpdate
      .Fields("LastUpdateByFK") = m_LastUpdateByID
      .Update
      m_ID = .Fields("OrgPK")
    End With
    rsOrg.Close
    Set rsOrg = Nothing
    m_IsValid = True
    Update = blnUpdate
  End Function

  Public Function Delete() ' As Boolean
    Dim blnDelete
    blnDelete = True
    m_IsActive = False
    If m_IsValid Then
      m_cnnDB.Execute "UPDATE tblOrganisations SET IsActive = " & m_IsActive & _
        " WHERE OrgPK = " & m_ID, , adExecuteNoRecords
    End If
    Delete = blnDelete
  End Function

  Public Function CheckPassword(ByVal f_Password) ' As Boolean
    ' #######
    ' Encrypt password in MD5
    CheckPassword = True
  End Function

  ' ###############
  ' Private Methods
  ' ###############
  Private Sub Class_Initialize()
    If IsObject(cnnDB) Then
      Set m_cnnDB = cnnDB
    End If
    m_IsValid = False

    ' Set Default Org Properties
    m_IsActive = True
    m_OrgShortName = "Unknown"
    m_OrgName = "Unknown"
  End Sub

  Private Sub Class_Terminate()
  End Sub

End Class
%>