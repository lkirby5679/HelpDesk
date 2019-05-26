<SCRIPT LANGUAGE=VBScript RUNAT=SERVER>

' ----------------------------------------------------------------------------------
'  Transworld Interactive Help Desk, Copyright (C) 1995-2011 Tom Kirby
'  Transworld Interactive Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  Filename: Settings.asp
'  Date:     $Date: 2004/03/11 06:08:42 $
'  Version:  $Revision: 1.3 $
'  Purpose:  Used the set the basic system parameters include database variables
' ----------------------------------------------------------------------------------

Sub SetAppVariables


	'==========================================================
	' Web Server Information
	'==========================================================
	
	Application("BASE_URL") = "http://support.transworldinteractive.net"   ' Don't end the string with a '/'  
	

	'==========================================================
	' Database Information
	'==========================================================

	' Database Type:
	'     1 - SQL Server with SQL security (set SQLUser/SQLPass)
	'     2 - SQL Server with integrated security
  '     3 - Access Database (set AccessPath)
	'     4 - DSN (An ODBC DataSource) (set DSN_Name)

	Application("DBType") = 3


	' SQL Database Settings
	
	Application("SQLServer") = "AA000891"	' Server name (don't put the leading \\)
	Application("SQLDBase") = "HURON"	    ' Database name
	Application("SQLUser") = "sa"			    ' Account to log into the SQL server with
	Application("SQLPass") = "password"		' Password for account


	' Access Database Settings
	
	Application("AccessPath") = "C:\Inetpub\wwwroot\twsupport\db\Huron.mdb"


	' DSN Link Settings
	
	Application("DSN_Name") = "HelpDeskDSN"



	'==========================================================
	' Debugging Information
	'==========================================================

	' Set to true to view full MS errors and other debug information printed, This
	' will disable most On Error Resume Next statements.
	
	Application("Debug") = False


End Sub


</SCRIPT>
