<%
' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: clsParameter.asp
'  Date:     $Date: 2004/03/11 06:29:08 $
'  Version:  $Revision: 1.2 $
'  Purpose:  Class object to assist when working with Parameters
' ----------------------------------------------------------------------------------

Class clsParameter
  ' ##################
  ' Private Properties
  ' ##################
  Private m_cnnDB

  ' #################
  ' Public Properties
  ' #################

  ' ##############
  ' Public Methods
  ' ##############
  Public Function GetValue(f_ParamName)  ' As String
    Dim strQuery, rsParam
      strQuery = "SELECT ParamValue FROM tblParameters WHERE ParamName = '" & f_ParamName & "'"
      Set rsParam = Server.CreateObject("ADODB.RecordSet")
      rsParam.Open strQuery, m_cnnDB
      If Not rsParam.EOF Then
        GetValue = rsParam("ParamValue")
      End If
      rsParam.Close
      Set rsParam = Nothing
  End Function

  Public Sub SetValue(f_ParamName,ByVal f_ParamValue)
    f_ParamValue = Left(Trim(f_ParamValue), 128)
    m_cnnDB.Execute "UPDATE tblParameters SET ParamValue = '" & f_ParamValue & "'" & _
      " WHERE ParamName = '" & f_ParamName & "'", , adExecuteNoRecords
  End Sub

  ' ###############
  ' Private Methods
  ' ###############
  Private Sub Class_Initialize()
    If IsObject(cnnDB) Then
      Set m_cnnDB = cnnDB
    End If
  End Sub

  Private Sub Class_Terminate()
  End Sub

End Class
%>