<%

' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: clsCaseType.asp
'  Date:     $Date: 2004/03/11 06:29:08 $
'  Version:  $Revision: 1.2 $
'  Purpose:  Class object to assist when working with Case Types
' ----------------------------------------------------------------------------------

Class clsCaseType


	' ----------------------------------------------------------------------------
	' Private Declarations

	Private m_ID, m_IsActive, m_IsValid, m_LastError, m_cnnDB
	Private m_CaseTypeName, m_CaseTypeDesc, m_CaseTypeOrder, m_RepGroup, m_RepGroupID
	Private m_LastUpdate, m_LastUpdateBy, m_LastUpdateByID


	' ----------------------------------------------------------------------------
	' Public Declarations

	Public SourceRS ' As RecordSet


	' ----------------------------------------------------------------------------
	' Public Properties

	Public Property Get ID()  ' As Long
		ID = m_ID
	End Property

	Public Property Let ID(f_ID)
		If IsNumeric(f_ID) Then
			m_ID = f_ID
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

	Public Property Get CaseTypeName()  ' As String
		CaseTypeName = m_CaseTypeName
	End Property

	Public Property Let CaseTypeName(ByRef f_CaseTypeName)
		m_CaseTypeName = Left(Trim(f_CaseTypeName), 36)
	End Property

	Public Property Get CaseTypeDesc()  ' As String
		CaseTypeDesc = m_CaseTypeDesc
	End Property

	Public Property Let CaseTypeDesc(ByRef f_CaseTypeDesc)
		m_CaseTypeDesc = Left(Trim(f_CaseTypeDesc), 160)
	End Property

	Public Property Get CaseTypeOrder()  ' As String
		CaseTypeOrder = m_CaseTypeOrder
	End Property

	Public Property Let CaseTypeOrder(ByRef f_CaseTypeOrder)
		m_CaseTypeOrder = f_CaseTypeOrder
	End Property

	Public Property Get RepGroup() ' As clsGroup
			  
		Dim objGroup
				  
		If Not IsObject(m_RepGroup) Then
				    
			If IsNumeric(m_RepGroupID) Then
					      
				Set objGroup = New clsGroup
						        
				objGroup.ID = m_RepGroupID

				If Not objGroup.Load Then
					' Raise error
				Else
'					Set m_RepGroup = objGroup
				End If
						        
				Set m_RepGroup = objGroup
						        
				Set objGroup = Nothing
					        
			Else
					      
				' Invalid RepGroup ID, Raise error
					      
			End If
				      
		Else
				    
			' m_RepGroup already loaded
				      
		End If
				    
		Set RepGroup = m_RepGroup
			    
	End Property

	Public Property Get RepGroupID() ' As Long
		RepGroupID = m_RepGroupID
	End Property

	Public Property Let RepGroupID(ByRef f_RepGroupID)

		Dim objGroup
				    
		If IsNumeric(f_RepGroupID) And Not IsEmpty(f_RepGroupID) And f_RepGroupID > 0 Then
				    
			Set objGroup = New clsGroup

			objGroup.ID = f_RepGroupID

			If Not objGroup.Load Then
				' Item not found, Raise Error
			Else
				m_RepGroupID = f_RepGroupID
				Set m_RepGroup = objGroup
			End If
					      
			Set objGroup = Nothing

		Else
				    
			' Invalid ID, Raise Error
				    
		End If
			    
	End Property


	Public Property Get IsValid() ' As Boolean
		IsValid = m_IsValid
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

	Public Property Get LastError() ' As String
		LastError = m_LastError
	End Property


	' ----------------------------------------------------------------------------
	' Private Properties



	' ----------------------------------------------------------------------------
	' Public Methods

	Public Function Load()  ' As Boolean

		' One of the following is required to load: (a) SourceRS, or (b) ID

		Dim blnLoad, rstCaseType, strQuery, blnUseSourceRS
		
		blnLoad = True
		blnUseSourceRS = False

		If IsObject(SourceRS) Then
			If SourceRS.State <> adStateClosed Then
				Set rstCaseType = SourceRS
				blnUseSourceRS = True
			Else
				blnLoad = False
				m_LastError = "SourceRS is closed."
			End If
		Else
			If IsNumeric(m_ID) Then
				strQuery = "SELECT * FROM tblCaseTypes WHERE CaseTypePK = " & m_ID
				Set rstCaseType = Server.CreateObject("ADODB.RecordSet")
				rstCaseType.Open strQuery, m_cnnDB
			Else
				blnLoad = False
				m_LastError = "Missing ID."
			End If
		End If

		If blnLoad Then
			If Not rstCaseType.EOF And Not rstCaseType.BOF Then
				
				On Error Resume Next
				
				m_ID = rstCaseType("CaseTypePK")
				m_IsActive = rstCaseType("IsActive")
				m_CaseTypeName = rstCaseType("CaseTypeName")
				m_CaseTypeDesc = rstCaseType("CaseTypeDesc")
				m_CaseTypeOrder = rstCaseType("CaseTypeOrder")
				m_RepGroupID = rstCaseType("RepGroupFK")
				m_LastUpdate = rstCaseType("LastUpdate")
				m_LastUpdateBy = rstCaseType("LastUpdateByFK")
				
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
		
		If Not blnUseSourceRS Then
			rstCaseType.Close
			Set rstCaseType = Nothing
		End If
		
		If blnLoad Then
			m_IsValid = True
		End If
		
		Load = blnLoad
		
	End Function

	Public Function Update() ' As Boolean
	
		Dim blnUpdate, rstCaseType
		
		blnUpdate = True
		
		Set rstCaseType = Server.CreateObject("ADODB.RecordSet")
		If m_IsValid Then
			rstCaseType.Open "SELECT * FROM tblCaseTypes WHERE CaseTypePK = " & m_ID, m_cnnDB, _
			adOpenKeyset, adLockOptimistic, adCmdText
		Else
			rstCaseType.Open "tblCaseTypes", m_cnnDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect
			rstCaseType.AddNew
		End If

		With rstCaseType
			.Fields("IsActive") = m_IsActive
			.Fields("CaseTypeName") = m_CaseTypeName
			.Fields("CaseTypeDesc") = m_CaseTypeDesc
			.Fields("CaseTypeOrder") = m_CaseTypeOrder
			.Fields("RepGroupFK") = m_RepGroupID
			.Fields("LastUpdate") = m_LastUpdate
			.Fields("LastUpdateByFK") = m_LastUpdateByID
			.Update
			m_ID = .Fields("CaseTypePK")
		End With

		rstCaseType.Close
		Set rstCaseType = Nothing

		m_IsValid = True
		Update = blnUpdate
		
	End Function

	Public Function Delete() ' As Boolean

		Dim blnDelete
		
		blnDelete = True
		m_IsActive = False
		
		If m_IsValid Then
			m_cnnDB.Execute "UPDATE tblCaseTypes SET IsActive = " & m_IsActive & _
							" WHERE CaseTypePK = " & m_ID, , adExecuteNoRecords
		End If
		
		Delete = blnDelete
		
	End Function


	' ----------------------------------------------------------------------------
	' Private Methods

	Private Sub Class_Initialize()
	
		If IsObject(cnnDB) Then
			Set m_cnnDB = cnnDB
		End If
		m_IsValid = False

		' Set Default Contact Properties
		m_IsActive = True
		m_CaseTypeName = "Unknown"
		
	End Sub

	Private Sub Class_Terminate()

	End Sub
	

End Class
%>