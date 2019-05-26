<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: clsCollection.asp
'  Date:     $Date: 2004/03/17 00:06:45 $
'  Version:  $Revision: 1.4 $
'  Purpose:  Class object to assist when working with collections of other classes
' ----------------------------------------------------------------------------------

Class clsCollection
  ' ##################
  ' Private Properties
  ' ##################
  Private m_cnnDB, m_Type
  Private m_Query, m_Item, m_ItemIsValid

  Private m_RecordSet
  Private m_LastError

  ' ##################
  ' Public Properties
  ' ##################
  Public clOrganisation, clContact, clDepartment, clGroup, clCase, clParameter
  Public clEmailMsg, clLanguage, clLanguageText 
  Public clPermission, clRole, clNote, clFile, clAssignment, clKnowledgebase
  Public clListItem
  Public clCaseType, clCategory

  Public Default Property Get Item()
    If m_ItemIsValid Then
      Set Item = m_Item
    Else
      Set m_Item = Nothing
      Dim blnTemp
      m_ItemIsValid = True
      Select Case m_Type
        Case clOrganisation
          Set m_Item = New clsOrganisation
          Set m_Item.SourceRS = m_RecordSet
          blnTemp = m_Item.Load
        Case 2
          Set m_Item = New clsContact
          Set m_Item.SourceRS = m_RecordSet
          blnTemp = m_Item.Load
        Case clDepartment
          Set m_Item = New clsDepartment
          Set m_Item.SourceRS = m_RecordSet
          blnTemp = m_Item.Load
        Case clGroup
          Set m_Item = New clsGroup
          Set m_Item.SourceRS = m_RecordSet
          blnTemp = m_Item.Load
        Case clCase
          Set m_Item = New clsCase
          Set m_Item.SourceRS = m_RecordSet
          blnTemp = m_Item.Load
        Case clParameter
          Set m_Item = New clsParameter
          Set m_Item.SourceRS = m_RecordSet
          blnTemp = m_Item.Load
        Case clEmailMsg
          Set m_Item = New clsEmailMsg
          Set m_Item.SourceRS = m_RecordSet
          blnTemp = m_Item.Load
        Case clLanguage
          Set m_Item = New clsLanguage
          Set m_Item.SourceRS = m_RecordSet
          blnTemp = m_Item.Load
        Case clLanguageText
          Set m_Item = New clsLanguageText
          Set m_Item.SourceRS = m_RecordSet
          blnTemp = m_Item.Load
        Case clPermission
          Set m_Item = New clsPerm
          Set m_Item.SourceRS = m_RecordSet
          blnTemp = m_Item.Load
        Case clRole
          Set m_Item = New clsRole
          Set m_Item.SourceRS = m_RecordSet
          blnTemp = m_Item.Load
        Case clNote
          Set m_Item = New clsNote
          Set m_Item.SourceRS = m_RecordSet
          blnTemp = m_Item.Load
        Case clFile
          Set m_Item = New clsFile
          Set m_Item.SourceRS = m_RecordSet
          blnTemp = m_Item.Load
        Case clListItem
          Set m_Item = New clsListItem
          Set m_Item.SourceRS = m_RecordSet
          blnTemp = m_Item.Load
        Case clCaseType
          Set m_Item = New clsCaseType
          Set m_Item.SourceRS = m_RecordSet
          blnTemp = m_Item.Load
        Case clCategory
          Set m_Item = New clsCategory
          Set m_Item.SourceRS = m_RecordSet
          blnTemp = m_Item.Load
        Case clAssignment
          Set m_Item = New clsAssignment
          Set m_Item.SourceRS = m_RecordSet
          blnTemp = m_Item.Load
        Case clKnowledgebase
          Set m_Item = New clsKnowledgebase
          Set m_Item.SourceRS = m_RecordSet
          blnTemp = m_Item.Load
        Case Else
          Set Item = Nothing
          m_ItemIsValid = False
      End Select
      Set Item = m_Item
    End If
  End Property

  Public Property Let CollectionType(f_Type)
    m_Type = f_Type
  End Property

  Public Property Get RecordCount() ' As Long
    RecordCount = m_RecordSet.RecordCount
  End Property

  Public Property Get BOF() ' As Boolean
    BOF = m_RecordSet.BOF
  End Property

  Public Property Get EOF() ' As Boolean
    EOF = m_RecordSet.EOF
  End Property

  Public Property Get PageCount() ' As Long
    PageCount = m_RecordSet.PageCount
  End Property

  Public Property Get PageSize()  ' As Long
    PageSize = m_RecordSet.PageSize
  End Property

  Public Property Let PageSize(f_PageSize)
    m_RecordSet.PageSize = f_PageSize
  End Property

  Public Property Get AbsolutePage()  ' As Long
    AbsolutePage = m_RecordSet.AbsolutePage
  End Property

  Public Property Let AbsolutePage(f_AbsolutePage)
    m_RecordSet.AbsolutePage = f_AbsolutePage
  End Property

  Public Property Get LastError() ' As String
    LastError = m_LastError
  End Property

  Public Property Let Query(ByRef f_Query)
    m_Query = Trim(f_Query)
  End Property

  ' ##################
  ' Public Methods
  ' ##################
  Public Function Load()  ' As Boolean
    Load = True
    If Len(m_Query) = 0 Then
      Load = False
      m_LastError = "Missing query string."
      Exit Function
    End If
    If Not IsNumeric(m_Type) Then
      Load = False
      m_LastError = "Invalid collection type."
      Exit Function
    End If
    m_RecordSet.CacheSize = 10
    m_RecordSet.Open m_Query, m_cnnDB, adOpenKeyset, adLockReadOnly
    m_ItemIsValid = False
  End Function

  Public Sub MoveNext()
    m_ItemIsValid = False
    m_RecordSet.MoveNext
  End Sub

  Public Sub MovePrevious()
    m_ItemIsValid = False
    m_RecordSet.MovePrevious
  End Sub

  Public Sub MoveFirst()
    m_ItemIsValid = False
    m_RecordSet.MoveFirst
  End Sub

  Public Sub MoveLast()
    m_ItemIsValid = False
    m_RecordSet.MoveLast
  End Sub

  Public Sub Move(f_Index)
    m_ItemIsValid = False
    m_RecordSet.Move(f_Index)
  End Sub
  ' ##################
  ' Private Methods
  ' ##################

  Private Sub Class_Initialize()
    If IsObject(cnnDB) Then
      Set m_cnnDB = cnnDB
    End If
    Set m_RecordSet = Server.CreateObject("ADODB.RecordSet")
    m_ItemIsValid = False
    clOrganisation = 1
    clContact = 2
    clDepartment = 3
    clGroup = 4
    clCase = 5
    clParameter = 6
    clEmailMsg = 7
    clLanguage = 8
    clPermission = 9
    clRole = 10
    clNote = 11
    clFile = 12
    clListItem = 13
    clCaseType = 14
    clCategory = 15
    clAssignment = 16
    clKnowledgebase = 17
    clLanguageText = 18
  End Sub

  Private Sub Class_Terminate()
    If m_RecordSet.State <> adStateClosed Then
      m_RecordSet.Close
    End If
    Set m_cnnDB = Nothing
    Set m_RecordSet = Nothing
  End Sub

End Class
%>