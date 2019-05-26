<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: clsEmailMsg.asp
'  Date:     $Date: 2004/03/11 06:29:08 $
'  Version:  $Revision: 1.2 $
'  Purpose:  Class object to assist when working with Email Message Types
' ----------------------------------------------------------------------------------

Class clsEMailMsg

	' ----------------------------------------------------------------------------
	' Private Declarations

	Private m_ID, m_IsActive, m_IsValid, m_LastError, m_cnnDB
	Private m_EMailMsgType, m_Subject, m_Body, m_Lang, m_LangID
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

	Public Property Get EMailMsgType()  ' As String
		EMailMsgType = m_EMailMsgType
	End Property

	Public Property Let EMailMsgType(ByRef f_EMailMsgType)	' As String
		m_EMailMsgType = f_EMailMsgType
	End Property

	Public Property Get Subject()  ' As String
		Subject = m_Subject
	End Property

	Public Property Let Subject(ByRef f_Subject)	' As String
		m_Subject = Left(Trim(f_Subject), 128)
	End Property

	Public Property Get Body()  ' As String
		Body = m_Body
	End Property

	Public Property Let Body(ByRef f_Body)  ' As String
		m_Body = f_Body
	End Property

	Public Property Get Lang()  ' As clsLanguage
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

		' One of the following is required to load: (a) SourceRS, or (b) ID, or (c) EMailMsgType

		Dim blnLoad, rstEMailMsg, strQuery, blnUseSourceRS
		
		blnLoad = True
		blnUseSourceRS = False

		If IsObject(SourceRS) Then
		
			If SourceRS.State <> adStateClosed Then
				Set rstEMailMsg = SourceRS
				blnUseSourceRS = True
			Else
				blnLoad = False
				m_LastError = "SourceRS is closed."
			End If
			
		Else
		
		    If Len(m_EMailMsgType) > 0 Then
				strQuery = "SELECT * from tblEMailMsgs WHERE EMailMsgType = '" & m_EMailMsgType & "'"
			ElseIf IsNumeric(m_ID) Then
				strQuery = "SELECT * FROM tblEMailMsgs WHERE EMailMsgPK = " & m_ID
			Else
				strQuery = ""
			End If
			
			If Len(strQuery) > 0 Then
				Set rstEMailMsg = Server.CreateObject("ADODB.RecordSet")
				rstEMailMsg.Open strQuery, m_cnnDB
			Else
				blnLoad = False
				m_LastError = "Missing ID."
			End If
			
		End If

		If blnLoad Then
			If Not rstEMailMsg.EOF And Not rstEMailMsg.BOF Then
				
				On Error Resume Next
				
				m_ID = rstEMailMsg("EMailMsgPK")
				m_IsActive = rstEMailMsg("IsActive")
				m_EMailMsgType = rstEMailMsg("EMailMsgType")
				m_Subject = rstEMailMsg("Subject")
				m_Body = rstEMailMsg("Body")
				m_LangID = rstEMailMsg("LangFK")
				m_LastUpdate = rstEMailMsg("LastUpdate")
				m_LastUpdateBy = rstEMailMsg("LastUpdateByFK")
				
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
			rstEMailMsg.Close
			Set rstEMailMsg = Nothing
		End If
		
		If blnLoad Then
			m_IsValid = True
		End If
		
		Load = blnLoad
		
	End Function

	Public Function Update() ' As Boolean
	
		Dim blnUpdate, rstEMailMsg
		
		blnUpdate = True
		
		Set rstEMailMsg = Server.CreateObject("ADODB.RecordSet")
		If m_IsValid Then
			rstEMailMsg.Open "SELECT * FROM tblEMailMsgs WHERE EMailMsgPK = " & m_ID, m_cnnDB, _
			adOpenKeyset, adLockOptimistic, adCmdText
		Else
			rstEMailMsg.Open "tblEMailMsgs", m_cnnDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect
			rstEMailMsg.AddNew
		End If

		With rstEMailMsg
			.Fields("IsActive") = m_IsActive
			.Fields("EMailMsgType") = m_EMailMsgType
			.Fields("Subject") = m_Subject
			.Fields("Body") = m_Body
			.Fields("LangFK") = m_LangID
			.Fields("LastUpdate") = m_LastUpdate
			.Fields("LastUpdateByFK") = m_LastUpdateByID
			.Update
			m_ID = .Fields("EMailMsgPK")
		End With

		rstEMailMsg.Close
		Set rstEMailMsg = Nothing

		m_IsValid = True
		Update = blnUpdate
		
	End Function

	Public Function Delete() ' As Boolean

		Dim blnDelete
		
		blnDelete = True
		m_IsActive = False
		
		If m_IsValid Then
			m_cnnDB.Execute "UPDATE tblEMailMsgs SET IsActive = " & m_IsActive & _
							" WHERE EMailMsgPK = " & m_ID, , adExecuteNoRecords
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
'		m_EMailMsgType = "Unknown"
		
	End Sub

	Private Sub Class_Terminate()

	End Sub
	

End Class
%>