<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: clsAssignment.asp
'  Date:     $Date: 2004/03/11 06:19:33 $
'  Version:  $Revision: 1.2 $
'  Purpose:  Object to assst when working with Assignments
' ----------------------------------------------------------------------------------

Class clsAssignment


	' ----------------------------------------------------------------------------
	' Private Declarations

	Private m_ID, m_IsActive, m_IsValid, m_LastError, m_cnnDB
	Private m_Rep, m_RepID, m_CaseType, m_CaseTypeID, m_Cat, m_CatID
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


	Public Property Get Rep() ' As clsContact

		Dim oRep, blnTemp


		If Not IsObject(m_Rep) Then
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
		    
			' Invalid ID, Raise Error
		    
		End If

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

		Dim blnLoad, rstAssignment, strQuery, blnUseSourceRS
		
		blnLoad = True
		blnUseSourceRS = False

		If IsObject(SourceRS) Then
			If SourceRS.State <> adStateClosed Then
				Set rstAssignment = SourceRS
				blnUseSourceRS = True
			Else
				blnLoad = False
				m_LastError = "SourceRS is closed."
			End If
		Else
			If IsNumeric(m_ID) Then
				strQuery = "SELECT * FROM tblAssignments WHERE AssignmentPK = " & m_ID
				Set rstAssignment = Server.CreateObject("ADODB.RecordSet")
				rstAssignment.Open strQuery, m_cnnDB
			Else
				blnLoad = False
				m_LastError = "Missing ID."
			End If
		End If

		If blnLoad Then
			If Not rstAssignment.EOF And Not rstAssignment.BOF Then
				
				On Error Resume Next
				
				m_ID = rstAssignment("AssignmentPK")
				m_IsActive = rstAssignment("IsActive")
				m_CaseTypeID = rstAssignment("CaseTypeFK")
				m_CatID = rstAssignment("CatFK")
				m_RepID = rstAssignment("RepFK")
				m_LastUpdate = rstAssignment("LastUpdate")
				m_LastUpdateBy = rstAssignment("LastUpdateByFK")
				
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
			rstAssignment.Close
			Set rstAssignment = Nothing
		End If
		
		If blnLoad Then
			m_IsValid = True
		End If
		
		Load = blnLoad
		
	End Function


	Public Function Update() ' As Boolean
	
		Dim blnUpdate, rstAssignment
		
		blnUpdate = True
		
		Set rstAssignment = Server.CreateObject("ADODB.RecordSet")
		If m_IsValid Then
			rstAssignment.Open "SELECT * FROM tblAssignments WHERE AssignmentPK = " & m_ID, m_cnnDB, _
			adOpenKeyset, adLockOptimistic, adCmdText
		Else
			rstAssignment.Open "tblAssignments", m_cnnDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect
			rstAssignment.AddNew
		End If

		With rstAssignment
			.Fields("IsActive") = m_IsActive
			.Fields("CaseTypeFK") = m_CaseTypeID
			.Fields("CatFK") = m_CatID
			.Fields("RepFK") = m_RepID
			.Fields("LastUpdate") = m_LastUpdate
			.Fields("LastUpdateByFK") = m_LastUpdateByID
			.Update
			m_ID = .Fields("AssignmentPK")
		End With

		rstAssignment.Close
		Set rstAssignment = Nothing

		m_IsValid = True
		Update = blnUpdate
		
	End Function


	Public Function Delete() ' As Boolean

		Dim blnDelete
		
		blnDelete = True
		m_IsActive = False
		
		If m_IsValid Then
			m_cnnDB.Execute "UPDATE tblAssignments SET IsActive = " & m_IsActive & _
							" WHERE AssignmentPK = " & m_ID, , adExecuteNoRecords
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
		
	End Sub


	Private Sub Class_Terminate()

	End Sub
	

End Class
%>