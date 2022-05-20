<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz Forums(TM)
'**  http://www.webwizforums.com
'**                            
'**  Copyright (C)2001-2018 Web Wiz Ltd. All Rights Reserved.
'**  
'**  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS UNDER LICENSE FROM WEB WIZ LTD.
'**  
'**  IF YOU DO NOT AGREE TO THE LICENSE AGREEMENT THEN WEB WIZ LTD. IS UNWILLING TO LICENSE 
'**  THE SOFTWARE TO YOU, AND YOU SHOULD DESTROY ALL COPIES YOU HOLD OF 'WEB WIZ' SOFTWARE
'**  AND DERIVATIVE WORKS IMMEDIATELY.
'**  
'**  If you have not received a copy of the license with this work then a copy of the latest
'**  license contract can be found at:-
'**
'**  https://www.webwiz.net/license
'**
'**  For more information about this software and for licensing information please contact
'**  'Web Wiz' at the address and website below:-
'**
'**  Web Wiz Ltd, Unit 18, The Glenmore Centre, Fancy Road, Poole, Dorset, BH12 4FB, England
'**  https://www.webwiz.net
'**
'**  Removal or modification of this copyright notice will violate the license contract.
'**
'****************************************************************************************





'Dimension global variables
Dim adoCon 			'Database Connection Variable Object
Dim strCon			'Holds the string to connect to the db
Dim rsCommon			'Holds the configuartion recordset
Dim strSQL			'Holds the SQL query for the database
Dim strDbPathAndName		'Holds the path and name of the database
Dim strSQLServerName		'Holds the name of the SQL Server
Dim strSQLDBUserName		'Holds the user name (for SQL Server Authentication)
Dim strSQLDBPassword		'Holds the password (for SQL Server Authentication)
Dim strSQLDBName		'Holds name of a database on the server
Dim strDatabaseDateFunction	'Holds a different date function for Access or SQL server
Dim strDatabaseType		'Holds the type of database used
Dim strDBFalse			'Holds the false value for SQL queries
Dim strDBTrue			'Holds the true value for SQL queries
Dim strDBNoLock			'Holds if the database is locked while running the query for SQL Server
Dim strRowLock			'Holds if the database row is locked while running the query for SQL Server
Dim strDBTop1			'Holds the SQL limit operator (TOP 1) for SQL Server and Access
Dim strDBLimit1			'Holds the SQL limit operator (LIMIT 1) for mySQL
Dim strMyODBCDriver		'MyODBC Driver for mySQL
Dim strAccessDriver		'Access DB Driver
Dim strSQLServersDriver 	'SQL Server Driver
Dim blnSqlSvrAdvPaging		'Set to true if advanced paging is used
Dim strDbTable			'Holds the database table prefix
Dim strDBO			'Holds the DBO Owner



'******************************************
'***   Database Connection settings    ****
'******************************************
%><!-- #include file="database_settings.asp" --><%



'******************************************
'*** 	 Open Database Connection      ****
'******************************************

'This sub procedure opens a connection to the database and creates a recordset object and sets database defaults
Public Sub openDatabase(strCon)

	'Setup database driver and defaults
	'**********************************
	
	'SQL Server Database Defaults
	If strDatabaseType = "SQLServer" Then
		
		'Please note this application has been optimised for the SQL OLE DB Driver using another driver 
		'or system DSN to connect to the SQL Server database will course errors in the application and
		'drastically reduce the performance!
		
		'The SQLOLEDB driver offers the highest performance at this time for connecting to SQL Server databases from within ASP.
		
		'MS SQL Server OLE Driver (If you change this string make sure you also change it in the msSQL_server_setup.asp file when creating the database)
		'strCon = "Provider=SQL Server Native Client 11.0;Connection Timeout=90;Command Timeout=200;Server=" & strSQLServerName & ";User ID=" & strSQLDBUserName & ";Password=" & strSQLDBPassword & ";Database=" & strSQLDBName & ";"
		strCon = "Provider=" & strSQLServersDriver & ";Server=" & strSQLServerName & ";User ID=" & strSQLDBUserName & ";Password=" & strSQLDBPassword & ";Database=" & strSQLDBName & ";"
	
		'The GetDate() function is used in SQL Server to get dates
		strDatabaseDateFunction = "GetDate()"
		
		'Set true and false for db
		strDBFalse = 0
		strDBTrue = 1
		
		'Set the lock variavbles for the db
		strDBNoLock = " WITH (NOLOCK) "
		strRowLock = " WITH (ROWLOCK) "
		
		'Set the Limit opertaor for SQL Server
		strDBTop1 = " TOP 1"
		
	
	'MySQL Server Database Defaults	
	ElseIf strDatabaseType = "mySQL" Then
		
		'This application requires the myODBC 3.51 or myODBC 5.1 driver
	
		'myODBC Driver
		strCon = "Driver={MySQL ODBC " & strMyODBCDriver & " Driver};Port=3306;Option=3;Server=" & strSQLServerName & ";User ID=" & strSQLDBUserName & ";Password=" & strSQLDBPassword & ";Database=" & strSQLDBName & ";"
		
		'Calculate the date web server time incase the database server is out, use international date
		strDatabaseDateFunction = "'" & internationalDateTime(Now())& "'"
		
		'Set true and false for db (true value is -1)
		strDBFalse = 0
		strDBTrue = -1
		
		
		'Set the limit operator
		strDBLimit1 = " LIMIT 1"
		
	
	'MS Access Database Defaults	
	ElseIf strDatabaseType = "Access" Then
		
	
		'Database driver
		strCon = "Provider=" & strAccessDriver & "; Data Source=" & strDbPathAndName
		
		'Database driver (Microsoft JET OLE DB driver version 4)
		'strCon = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & strDbPathAndName
		
		'Database driver (Microsoft ACE OLE DB driver) for Access 2007 and above and for x64
		'strCon = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & strDbPathAndName
		
		'The now() function is used in Access for dates
		strDatabaseDateFunction = "Now()"
		
		'Set true and false for db
		strDBFalse = "false"
		strDBTrue = "true"
			
		'Set the limit operator for Access
		strDBTop1 = " TOP 1"
		
	End If
	

	
	'Open Datbase Connection
	'***********************

	'Create a db connection odject
	Set adoCon = CreateObject("ADODB.Connection")
	
	'Set error trapping
	On Error Resume Next
	
	'Set the connection string to the database
	adoCon.connectionstring = strCon
	
	'Set an active connection to the Connection object
	adoCon.Open
	
	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then	Call errorMsg("An error has occurred while connecting to the database.", "db_connection", "database_connection.asp")
					
	'Disable error trapping
	On Error goto 0
	
	
	'Intialise the main ADO recordset object
	Set rsCommon = CreateObject("ADODB.Recordset")

End Sub




'******************************************
'*** 	  Close Database Connection    ****
'******************************************

'This sub procedure will close the main recordset and close the database connection
Public Sub closeDatabase()

	'Close recordset
	If isObject(rsCommon) Then
		Set rsCommon = Nothing
	End If
	
	'Close Database Connection
	If isObject(adoCon) Then
		adoCon.Close
		Set adoCon = Nothing
	End If
End Sub
%>