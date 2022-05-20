<% @ Language=VBScript %>
<% Option Explicit %>
<!-- #include file="database/database_connection.asp" -->
<!-- #include file="includes/version_inc.asp" -->
<!-- #include file="functions/functions_filters.asp" -->
<!-- #include file="functions/functions_common.asp" -->
<!-- #include file="functions/functions_report_errors.asp" -->
<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz Forums(TM)
'**  http://www.webwizforums.com
'**                            
'**  Copyright (C)2001-2019 Web Wiz Ltd. All Rights Reserved.
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




'Set the response buffer to true as we maybe redirecting and setting a cookie
Response.Buffer = True


'If the database is already setup then move 'em
If strDatabaseType <> "" Then Response.Redirect("default.asp")
	
'Dimension variables
Dim objFSO
Dim objFile
Dim strFileLocation
Dim strFileAspContents
Dim blnDetailedErrorReporting
Dim blnErrorLogging
Dim strLoggedInUsername
Dim blnLoggingEnabled
Dim strErrorMessage
Dim strAccessDbPath
Dim strAccessDbFileName
Dim strAccessOrgFileLocation
Dim blnAccessDbPhysicalPath


'Enable detailed error reporting
blnDetailedErrorReporting = True
blnErrorLogging = False
blnLoggingEnabled = False
blnAccessDbPhysicalPath = False
	
'File location
strAccessOrgFileLocation = Server.MapPath("database\wwForum.mdb")
strFileLocation = Server.MapPath("database\database_settings.asp")

'Set defaults
strDbPathAndName = "database/wwForum.mdb"
strMyODBCDriver = "3.51"
strDbTable = "tbl"
strDBO = "DBO"
strAccessDbPath = "database"
strAccessDbFileName = "wwForum.mdb"
strAccessDriver = "Microsoft.ACE.OLEDB.12.0"


'If a postback run the code
If Request.Form("postBack") Then
		
		
	'*** SQL Server ****
	If Request.Form("DbType") = "SQLServer" Then
		
		strDatabaseType = "SQLServer"
	
	'*** SQL Server 2000 ****	
	ElseIf Request.Form("DbType") = "SQLServer2000" Then
		
		strDatabaseType = "SQLServer"
	
	'*** mySQL ****	
	ElseIf Request.Form("DbType") = "mySQL" Then
		
		strDatabaseType = "mySQL"
	
	'*** Access ****		
	ElseIf Request.Form("DbType") = "Access" Then
		'Db Type
		strDatabaseType = "Access"
	End If

	'Read in form contents
	If strDatabaseType <> "Access" Then
		strSQLServerName = Trim(Mid(Request.Form("DbServer"), 1, 80))
		strSQLServerName = formatSQLInput(strSQLServerName)
		strSQLDBName = Trim(Mid(Request.Form("DbName"), 1, 50))
		strSQLDBName = formatSQLInput(strSQLDBName)
		strSQLDBUserName = Trim(Mid(Request.Form("DbUsername"), 1, 50))
		strSQLDBUserName = formatSQLInput(strSQLDBUserName)
		strSQLDBPassword = Trim(Mid(Request.Form("DbPassword"), 1, 50))
		strSQLDBPassword = formatSQLInput(strSQLDBPassword)
	
	'Else read in Access details
	Else
		strAccessDbPath = Trim(Mid(Request.Form("AccessDbPath"), 1, 80))
		strAccessDbFileName = Trim(Mid(Request.Form("AccessDbName"), 1, 20))
		blnAccessDbPhysicalPath = CBool(Request.Form("AccessDbServerPath"))
		
		'Check access DB path is OK
		strAccessDbPath = Replace(strAccessDbPath, "/", "\")
		If strAccessDbPath <> "" AND isNull(strAccessDbPath) = False Then
			If Mid(strAccessDbPath, len(strAccessDbPath), 1) <> "\" Then  strAccessDbPath = strAccessDbPath & "\"
		End If
			
	End If
	If Request.Form("DbOwner") <> "" Then strDBO = Trim(Mid(Request.Form("DbOwner"), 1, 50))
	If Request.Form("DbTablePrefix") <> "" Then strDbTable = Trim(Mid(Request.Form("DbTablePrefix"), 1, 50))
		
	strDBO = formatSQLInput(strDBO)
	'Set the myODBC driver
	If Request.Form("myODBC") = "5.3 ANSI" Then
		strMyODBCDriver = "5.3 ANSI"
	ElseIf Request.Form("myODBC") = "5.2 ANSI" Then
		strMyODBCDriver = "5.2 ANSI"
	ElseIf Request.Form("myODBC") = "5.1" Then
		strMyODBCDriver = "5.1"
	Else
		strMyODBCDriver = "3.51"
	End If
	'Set the paging type
	If Request.Form("DbType") = "SQLServer2000" Then
		blnSqlSvrAdvPaging = "False"
	Else
		blnSqlSvrAdvPaging = "True"
	End If	
	'Set SQL Server Driver
	If Request.Form("SQLDriver") = "SQLOLEDB" Then
		strSQLServersDriver = "SQLOLEDB"
	Else
		strSQLServersDriver = "SQL Server Native Client 11.0"
	End If
	'Set Access Driver
	If Request.Form("AccessDriver") = "JET" Then
		strAccessDriver = "Microsoft.Jet.OLEDB.4.0"
	Else
		strAccessDriver = "Microsoft.ACE.OLEDB.12.0"
	End If





	'SQL Server Connection String
	If strDatabaseType = "SQLServer" Then
				
		'MS SQL Server OLE Driver (If you change this string make sure you also change it in the msSQL_server_setup.asp file when creating the database)
		strCon = "Provider=" & strSQLServersDriver & ";Server=" & strSQLServerName & ";User ID=" & strSQLDBUserName & ";Password=" & strSQLDBPassword & ";Database=" & strSQLDBName & ";"
		
			
	'MySQL Server Connection String
	ElseIf strDatabaseType = "mySQL" Then
			
		'myODBC Driver
		strCon = "Driver={MySQL ODBC " & strMyODBCDriver & " Driver};Port=3306;Option=3;Server=" & strSQLServerName & ";User ID=" & strSQLDBUserName & ";Password=" & strSQLDBPassword & ";Database=" & strSQLDBName & ";"
			
			
	'MS Access Connection String
	ElseIf strDatabaseType = "Access" Then
		
		
		'Set error trapping
		On Error Resume Next
	
			
		'Location of database
		If blnAccessDbPhysicalPath Then
			strDbPathAndName = strAccessDbPath & strAccessDbFileName
		Else
			strDbPathAndName = Server.MapPath(strAccessDbPath & strAccessDbFileName)
		End If 
			
		'If an error has occurred write an error to the page
		If Err.Number <> 0 Then	strErrorMessage = ("<strong>An Error Occurred Connecting to the Database</strong><br /><br /><strong>Error details:-</strong><br />" & Err.Source & "<br />" & Err.Description & "<br /><br />")
							
		'Disable error trapping
		On Error goto 0
			
		'Database driver
		strCon = "Provider=" & strAccessDriver & "; Data Source=" & strDbPathAndName
	
		
		
		'Move database to new location (removed as due to support issues)
		'dim fs
		'set fs=Server.CreateObject("Scripting.FileSystemObject")
		'fs.CopyFile strAccessOrgFileLocation, Server.MapPath(strAccessDbPath & "/" & strAccessDbFileName)
		'set fs=nothing
			
	End If
	
	
	'Open Datbase Connection
	'***********************
	If strErrorMessage = "" Then
	
		'Create a db connection odject
		Set adoCon = CreateObject("ADODB.Connection")
			
		'Set error trapping
		On Error Resume Next
			
		'Set the connection string to the database
		adoCon.connectionstring = strCon
			
		'Set an active connection to the Connection object
		adoCon.Open
			
		'If an error has occurred write an error to the page
		If Err.Number <> 0 Then	strErrorMessage = ("<strong>An Error Occurred Connecting to the Database</strong><br /><br /><strong>Error details:-</strong><br />" & Err.Source & "<br />" & Err.Description & "<br /><br />")
							
		'Disable error trapping
		On Error goto 0
		
		'Clean up
		Set adoCon = Nothing
	End If

End If




'If no error has occured and this is a postback create the databasse setup file
If Request.Form("postBack") AND strErrorMessage = "" Then 	
	
	
	'If the db connection is good it will continue
	
	
	'Create the connection_settings.asp file contents
	strFileAspContents = "" & _
	VbCrLf & _
	VbCrLf & "'******************************************" & _
	VbCrLf & "'*** 	  Database System Type         ****" & _
	VbCrLf & "'******************************************" & _
	VbCrLf & _
	VbCrLf & "'Database Type"
	
	'*** SQL Server ***
	If Request.Form("DbType") = "SQLServer" OR Request.Form("DbType") = "SQLServer2000" Then
		
		strFileAspContents = strFileAspContents & "" & _
		VbCrLf & "strDatabaseType = ""SQLServer""	'Microsoft SQL Server 2000, 2005, 2008, 2008 R2 (Supports Enterprise, Standard, Workgroup, Web, and Express Editions)" & _
		VbCrLf & "'strDatabaseType = ""mySQL""	'MySQL 4.1 or MySQL 5.x" & _
		VbCrLf & "'strDatabaseType = ""Access""	'Microsoft Access Database (Very slow, not very good, best off avoided)"
		
	
	'*** mySQL ****	
	ElseIf Request.Form("DbType") = "mySQL" Then
		
		strFileAspContents = strFileAspContents & "" & _
		VbCrLf & "'strDatabaseType = ""SQLServer""	'Microsoft SQL Server 2000, 2005, 2008, 2008 R2 (Supports Enterprise, Standard, Workgroup, Web, and Express Editions)" & _
		VbCrLf & "strDatabaseType = ""mySQL""	'MySQL 4.1 or MySQL 5.x" & _
		VbCrLf & "'strDatabaseType = ""Access""	'Microsoft Access Database (Very slow, not very good, best off avoided)"
		
	
	'*** Access ****		
	ElseIf Request.Form("DbType") = "Access" Then
		
		strFileAspContents = strFileAspContents & "" & _
		VbCrLf & "'strDatabaseType = ""SQLServer""	'Microsoft SQL Server 2000, 2005, 2008, 2008 R2 (Supports Enterprise, Standard, Workgroup, Web, and Express Editions)" & _
		VbCrLf & "'strDatabaseType = ""mySQL""	'MySQL 4.1 or MySQL 5.x" & _
		VbCrLf & "strDatabaseType = ""Access""	'Microsoft Access Database (Very slow, not very good, best off avoided)"
		
	End If
	
	
	
		
		
	'Build the file contents
	strFileAspContents = strFileAspContents & "" & _
	VbCrLf & _
	VbCrLf & _
	VbCrLf & _
	VbCrLf & "'******************************************"  & _
	VbCrLf & "'*** 	      Microsoft Access         ****"  & _
	VbCrLf & "'******************************************"  & _
	VbCrLf & _
	VbCrLf & "'Microsoft Access is a flat file database system, it suffers from slow performance, limited "  & _
	VbCrLf & "'connections, and as a flat file it can be easly downloaded by a hacker if you do not secure "  & _
	VbCrLf & "'the database file!"  & _
	VbCrLf
	
	'Access Db Path (MapPath)
	If blnAccessDbPhysicalPath = False Then
		
		strFileAspContents = strFileAspContents & "" & _
		VbCrLf & "'Virtual path to database"  & _
		VbCrLf & "strDbPathAndName = Server.MapPath(""" & strAccessDbPath & strAccessDbFileName & """)  'This is the path of the database from the applications location"  & _
		VbCrLf & _	
		VbCrLf & "'Physical path to database"  & _
		VbCrLf & "'strDbPathAndName = """" 'Use this if you use the physical server path, eg:- ""C:\Inetpub\private\wwForum.mdb"""
	
	'Access Db Path (Physical)
	Else
		
		strFileAspContents = strFileAspContents & "" & _
		VbCrLf & "'Virtual path to database"  & _
		VbCrLf & "'strDbPathAndName = Server.MapPath(""database\wwForum.mdb"")  'This is the path of the database from the applications location"  & _
		VbCrLf & _	
		VbCrLf & "'Physical path to database"  & _
		VbCrLf & "strDbPathAndName = """ & strAccessDbPath & strAccessDbFileName & """ 'Use this if you use the physical server path, eg:- ""C:\Inetpub\private\wwForum.mdb"""
	End If
		
	
	'Build the file contents
	strFileAspContents = strFileAspContents & "" & _
	VbCrLf & _
	VbCrLf & _
	VbCrLf & "'PLEASE NOTE: - For extra security it is highly recommended you change the name of the database, wwForum.mdb, "  & _
	VbCrLf & "'to another name and then replace the wwForum.mdb found above with the name you changed the forum database to."  & _
	VbCrLf & _
	VbCrLf & _
	VbCrLf & _
	VbCrLf & "'**********************************************************"  & _
	VbCrLf & "'*** 	   Microsoft SQL Server and MySQL Server        ****"  & _
	VbCrLf & "'**********************************************************"  & _
	VbCrLf & _
	VbCrLf & "'Enter the details of your Microsoft SQL Server or MySQL Server and database below"  & _
	VbCrLf & "'*********************************************************************************"  & _
	VbCrLf & _	
	VbCrLf & "strSQLServerName = """ & strSQLServerName & """ 'Holds the name of the SQL Server (This is the name/location or IP address of the SQL Server)"  & _
	VbCrLf & "strSQLDBUserName = """ & strSQLDBUserName & """ 'Holds the user name (for SQL Server Authentication)"  & _
	VbCrLf & "strSQLDBPassword = """ & strSQLDBPassword & """ 'Holds the password (for SQL Server Authentication)"  & _
	VbCrLf & "strSQLDBName = """ & strSQLDBName & """"  & _
	VbCrLf & _
	VbCrLf & "'*** Advanced Paging - Performance Boost ***"  & _
	VbCrLf & "'Set this to true for advanced paging in SQL Server 2005/2008 and mySQL "  & _
	VbCrLf & "'If you use SQL Server 2005/2008 or mySQL this will give a massive performance boost to your forum"  & _
	VbCrLf & "blnSqlSvrAdvPaging = " & blnSqlSvrAdvPaging & "" & _
	VbCrLf & _	
	VbCrLf & _	
	VbCrLf & "'*** SQL Server Database Driver ***"  & _
	VbCrLf & "''Database Driver for SQL Server"  & _
	VbCrLf & "strSQLServersDriver = """ & strSQLServersDriver & """"  & _
	VbCrLf & _	
	VbCrLf & _	
	VbCrLf & "'*** SQL Server DBO Owner ***"  & _
	VbCrLf & "''Sets the schema owner for SQL Server (Usually DBO (DataBase Owner))"  & _
	VbCrLf & "strDBO = """ & strDBO & """"  & _
	VbCrLf & _	
	VbCrLf & _	
	VbCrLf & "'*** mySQL Database Driver ***"  & _
	VbCrLf & "'Web Wiz Forums supports myODBC 3.51 to myODBC 5.3 database drivers when used with the mySQL database. "  & _
	VbCrLf & "'Most web host support myODBC 3.51, but if your web host supports myODBC 5.1 I would recommend that you use that instead"  & _
	VbCrLf & "strMyODBCDriver = """ & strMyODBCDriver & """"  & _
	VbCrLf & _	
	VbCrLf & _	
	VbCrLf & "'*** Access Database Driver ***"  & _
	VbCrLf & "'Web Wiz NewsPad supports JET and ACE database drivers when used with the Access database. "  & _
	VbCrLf & "'If you are using an x64 application pool you would need to use the ACE driver"  & _
	VbCrLf & "strAccessDriver = """ & strAccessDriver & """"  & _
	VbCrLf & _
	VbCrLf & _
	VbCrLf & "'Set up the database table name prefix"  & _
	VbCrLf & "'(This is useful if you are running multiple forums from one database)"  & _
	VbCrLf & "strDbTable = """ & strDbTable & """"  & _
	VbCrLf & _
	VbCrLf	
		
		
		
	'Set error trapping
	On Error Resume Next
		
	'Creat an instance of the FSO object
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		
	'Check to see if an error has occurred
	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then	Call errorMsg("An error has occurred during logging.<br />Please check the File System Object (FSO) is installed on the server.", "create_FSO_object", "setup_database_connection.asp")
	
	'Disable error trapping
	On Error goto 0
	
	
	'Set error trapping
	On Error Resume Next		
			
	'See if the folder and file exist, if not create them
	Set objFile = objFSO.CreateTextFile(strFileLocation)
	
	'If an error has occurred writing the database_settings.asp file then redirect to permissions checking page
	If Err.Number <> 0 Then	
		
		'Close the file and clean up
		objFile.Close
		Set objFile = Nothing
		Set objFSO = Nothing
		
		If Err.Number <> 0 Then	strErrorMessage = ("<strong>An Error Occurred Writing the Database Settings File</strong><br /><br />Please check that you have Read, Write, and Modify Permissions on the Web Wiz Forums '<strong>Database</strong>' Directory.<br /><br />For further information on setting permissions take a look at the following knowledgebase article, <a href=""https://www.webwiz.net/kb/asp_knowledgebase/server_permissions.asp"" target=""_blank"">Checking and Setting up the Correct Permissions on the IIS Server</a>.<br /><br />")
	
	Else
	
		'Disable error trapping
		On Error goto 0		
			
				
		'Write to the a new line to the log file
		objFile.WriteLine(chr(60) & chr(37)) 'ASP open tag
		objFile.WriteLine(strFileAspContents)
		objFile.WriteLine(chr(37) & chr(62)) 'ASP close tag
			
			
		'Close the file and clean up
		objFile.Close
		Set objFile = Nothing
		Set objFSO = Nothing	
		
		
		
		'Redirect to database setup depending on what type of db setup we have
		
		'New install
		If Request.Form("install") = "new" Then
		
			If strDatabaseType = "SQLServer" Then
				Response.Redirect("setup_db.asp?setup=SqlServerNew")	
			ElseIf strDatabaseType = "mySQL" Then
				Response.Redirect("setup_db.asp?setup=mySQLNew")
			ElseIf strDatabaseType = "Access" Then	
				Response.Redirect("setup_db.asp?setup=AccessNew")
			End If
		
		'11x Update	
		ElseIf Request.Form("install") = "upgrade11" Then
			Response.Redirect("setup_db.asp?setup=11Update")
		
		'10x Update
		ElseIf Request.Form("install") = "upgrade10" Then
			
			If strDatabaseType = "SQLServer" Then
				Response.Redirect("setup_db.asp?setup=SqlServer10Update")	
			ElseIf strDatabaseType = "mySQL" Then
				Response.Redirect("setup_db.asp?setup=mySQL10Update")
			ElseIf strDatabaseType = "Access" Then	
				Response.Redirect("setup_db.asp?setup=Access10Update")
			End If

		
		'9x Update
		ElseIf Request.Form("install") = "upgrade9" Then
			
			If strDatabaseType = "SQLServer" Then
				Response.Redirect("setup_db.asp?setup=SqlServer9Update")	
			ElseIf strDatabaseType = "mySQL" Then
				Response.Redirect("setup_db.asp?setup=mySQL9Update")
			ElseIf strDatabaseType = "Access" Then	
				Response.Redirect("setup_db.asp?setup=Access9Update")
			End If
		
		'8x Update
		ElseIf Request.Form("install") = "upgrade8" Then
			
			If strDatabaseType = "SQLServer" Then
				Response.Redirect("setup_db.asp?setup=SqlServer8Update")	
			ElseIf strDatabaseType = "mySQL" Then
				Response.Redirect("setup_db.asp?setup=mySQL8Update")
			ElseIf strDatabaseType = "Access" Then	
				Response.Redirect("setup_db.asp?setup=Access8Update")
			End If
			
		'7x Update
		ElseIf Request.Form("install") = "upgrade7" Then
			
			If strDatabaseType = "SQLServer" Then
				Response.Redirect("setup_db.asp?setup=SqlServer7Update")	
			ElseIf strDatabaseType = "Access" Then	
				Response.Redirect("setup_db.asp?setup=Access7Update")
			End If
		
		Else
			Response.Redirect("setup_db.asp?setup=12Update")	
		
		
		End If
	End If

End If



%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en" dir="ltr">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Web Wiz Forums Configure Database Connection</title>
<meta name="generator" content="Web Wiz Forums" />
<%

Response.Write(vbCrLf  & vbCrLf & "<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Forums(TM) " & _
vbCrLf & "Info: http://www.webwizforums.com" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->" & vbCrLf)
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******

%>
<link href="css_styles/default/default_style.css" rel="stylesheet" type="text/css" />
<script  language="JavaScript" type="text/javascript">
	
//function to open pop up window
function winOpener(theURL, winName, scrollbars, resizable, width, height) {

	winFeatures = 'left=' + (screen.availWidth-10-width)/2 + ',top=' + (screen.availHeight-30-height)/2 + ',scrollbars=' + scrollbars + ',resizable=' + resizable + ',width=' + width + ',height=' + height + ',toolbar=0,location=0,status=1,menubar=0'
  	window.open(theURL, winName, winFeatures);
}
	
//Function to test email window
function OpenTestDbConnection(formName){

	now = new Date; 
	submitAction = formName.action;
	submitTarget = formName.target;
	
	//Open the window first 	
   	winOpener('','testDbCon',1,1,550,400)
   		
   	//Now submit form to the new window
   	formName.action = 'setup_db_con_test.asp?ID=' + now.getTime();	
	formName.target = 'testDbCon';
	formName.submit();
	
	//Reset submission
	formName.action = submitAction;
	formName.target = submitTarget;
}
</script>
</head>
<body>
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
  <tr>
    <td><a href="http://www.webwizforums.com"><img src="forum_images/web_wiz_forums.png" border="0" alt="Web Wiz Forums Homepage" title="Web Wiz Forums Homepage" /><br />
      <br />
      </a>
      <h1>Configure Database Connection</h1>
      <span class="smText"><br />
      You can configure the database settings used by Web Wiz Forums on this page. If you are installing Web Wiz Forums  in a 'Hosting Account' your hosting provider should have provided you with the information. </span></td>
  </tr>
</table>
<%

'If an error has occuered display the error message
If strErrorMessage <> "" Then

%>
<br />
<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
  <tr>
    <td><img src="forum_images/error.png" alt="Error" /> <strong>Error</strong></td>
  </tr>
  <tr>
    <td><%
	Response.Write(strErrorMessage)
  	%></td>
  </tr>
</table>
<%

End If

%>
<form id="frmSetup" name="frmSetup" method="post" action="">
  <br />
  <table cellspacing="1" cellpadding="3" class="tableBorder" align="center">
    <tr class="tableLedger">
      <td colspan="2">Database Connection Details</td>
    </tr>
    <tr>
      <td width="666" align="left"  class="tableRow">Database Type</td>
      <td width="919" valign="top"  class="tableRow"><strong>
        <%
'*** SQL Server 2005/2008 ****	    	
If Request.Form("DbType") = "SQLServer" Then
	Response.Write("Microsoft SQL Server")

'*** SQL Server 2000 ****	
ElseIf Request.Form("DbType") = "SQLServer2000" Then
	Response.Write("Microsoft SQL Server 2000")
'*** mySQL ****	
ElseIf Request.Form("DbType") = "mySQL" Then
	
	Response.Write("mySQL 4.1 or above")
'*** Access ****	
ElseIf Request.Form("DbType") = "Access" Then
	
	Response.Write("Microsoft Access")
End If

%>
        </strong></td>
    </tr>
    <%
 
'If access
If Request.Form("DbType") = "Access" Then
%>
    <tr>
      <td align="left" valign="top"  class="tableRow">Access Database File Name<br />
        <span class="smText">This is the name of your Access database (eg. Forum.mdb).</span></td>
      <td valign="top"  class="tableRow"><input name="AccessDbName" type="text" id="AccessDbName" value="<% = strAccessDbFileName %>" size="20" /></td>
    </tr>
    <tr>
      <td align="left" valign="top"  class="tableRow">Path to Access Database Folder<br />
        <span class="smText">This is the path from this web page to the location of the folder on the server where your Access database will be located. Make sure that the folder has<a href="https://www.webwiz.net/kb/asp_knowledgebase/server_permissions.asp" target="_blank" class="smLink"> Read, Write, and Modify Permissions set</a>.</span></td>
      <td valign="top"  class="tableRow"><input name="AccessDbPath" type="text" id="AccessDbPath" value="<% = strAccessDbPath %>" size="80" />
      	 <br />
      	 <input name="AccessDbServerPath" type="radio" value="false" checked="checked" />
    Path from this application to database
    <br />
     <input name="AccessDbServerPath" type="radio" value="true" />
    Physical Server Path to Database <span class="smText">(not URL) </span>
   
      </td>
    </tr>
     <tr>
      <td align="left"  class="tableRow">Access Database Driver:<br />
        <span class="style1"><span class="smText">This is the Access Driver version installed on the web server that is used to connect to your Access database</span></td>
      <td valign="top"  class="tableRow">
      	<label>
      	  <input name="AccessDriver" type="radio" id="AccessDriver" value="ACE"<% If strAccessDriver = "" OR strAccessDriver ="Microsoft.ACE.OLEDB.12.0" Then Response.Write(" checked=""checked""") %> />
          Microsoft ACE Access Driver
        </label>
        <br />
        <label>
          <input type="radio" name="AccessDriver" id="AccessDriver" value="JET"<% If strAccessDriver ="Microsoft.Jet.OLEDB.4.0" Then Response.Write(" checked=""checked""") %> />
          Microsoft JET Access Driver
        </label>
     </td>
    </tr>
    <%

'All other db types
Else
	
%>
    <tr>
      <td align="left"  class="tableRow">&nbsp;</td>
      <td valign="top"  class="tableRow">&nbsp;</td>
    </tr>
    <tr>
      <td align="left"  class="tableRow">Server<br />
        <span class="smText">This is the Host Name or IP Address of the Database Server</span></td>
      <td valign="top"  class="tableRow"><input name="DbServer" type="text" id="DbServer" value="<% = strSQLServerName %>" maxlength="80" /></td>
    </tr>
    <tr>
      <td align="left"  class="tableRow">Database Name<br />
        <span class="smText">This is the name of your Database</span></td>
      <td valign="top"  class="tableRow"><input name="DbName" type="text" id="DbName" value="<% = strSQLDBName %>" maxlength="50" /></td>
    </tr>
    <tr>
      <td align="left"  class="tableRow">Database Username<br />
        <span class="smText">This is your Database Login Name</span></td>
      <td valign="top"  class="tableRow"><input name="DbUsername" type="text" id="DbUsername" maxlength="50" /></td>
    </tr>
    <tr>
      <td align="left"  class="tableRow">Database Password<br />
        <span class="smText">This is your Database Password</span></td>
      <td valign="top"  class="tableRow"><input name="DbPassword" type="password" id="DbPassword" maxlength="50" /></td>
    </tr>
   <%
  
	If Request.Form("DbType") = "SQLServer" OR Request.Form("DbType") = "SQLServer2000"  Then
	
%>  
    
    <tr>
      <td align="left"  class="tableRow"><span class="tableRow">SQL Server  Database Driver:<br />
      <span class="style1"><span class="smText">This is the SQL Server Driver version installed on the web server that is used to connect to your SQL Server database</span></span></td>
      <td valign="top"  class="tableRow"><label>
        <input name="SQLDriver" type="radio" id="SQLDriver" value="SQLOLEDB"<% If strSQLServersDriver = "" OR strSQLServersDriver ="SQLOLEDB" Then Response.Write(" checked=""checked""") %> />
        SQL OLE DB (SQLOLEDB)</label><br />
        <label>
          <input type="radio" name="SQLDriver" id="SQLDriver" value="SQLNative"<% If strSQLServersDriver ="SQL Server Native Client 11.0" Then Response.Write(" checked=""checked""") %> />
          SQL Server Native Client 11.0</label></td>
    </tr>
    <tr>
      <td colspan="2" align="left"  class="tableSubLedger">Advanced Options</td>
    </tr>
   
    <tr>
      <td align="left"  class="tableRow">Object Owner<br />
        <span class="smText">This is the name of the group owning the objects (tables, index's, etc.) in the database. Most times this should be left as 'DBO'</span></td>
      <td valign="top"  class="tableRow"><input name="DbOwner" type="text" id="DbOwner" value="<% = strDBO %>" maxlength="50" />
        <br />
        (DBO - <strong>D</strong>ata<strong>B</strong>ase <strong>O</strong>wner)</td>
    </tr>
    <%
  
	End If

%>
    <tr>
      <td align="left"  class="tableRow">Database Table Prefix<br />
        <span class="smText"><span class="smText"><% 
        
        'Tell people not to change the database prefix
        If Request.Form("install") = "new" Then
        
        	%>This should be left as the default 'tbl' unless you have multiple installations of Web Wiz Forums sharing the same database, then each installation must have a different prefix.<%

	'Change the wording for upgrades otherwise you have people thinking that by using a new table prefix it will be able to locate the old tables (as they say, create something idiot proof and they make a better idiot!)
        Else
        	
        	%>If you changed the default Web Wiz Forums database table prefix from 'tbl' during the original install then you need to use the same database table prefix to upgrade your database. If you did not change the table prefix then leave it as 'tbl'.<%
        End If
        	
        	%></span></td>
      <td valign="top"  class="tableRow"><input name="DbTablePrefix" type="text" id="DbTablePrefix" value="<% = strDbTable %>" maxlength="25" /></td>
    </tr>
    <%

	If Request.Form("DbType") = "mySQL" Then
	
%>
    <tr>
      <td align="left"  class="tableRow">myODBC Database Driver:<br />
        <span class="style1"><span class="smText">This is the myODBC Driver version installed on the web server that is used to connect to your mySQL database</span></td>
      <td valign="top"  class="tableRow">
      	<label>
      	  <input name="myODBC" type="radio" id="myODBC1" value="3.51"<% If strMyODBCDriver = "" OR strMyODBCDriver ="3.51" Then Response.Write(" checked=""checked""") %> />
          myODBC  3.51
        </label>
        <br />
        <label>
          <input type="radio" name="myODBC" id="myODBC2" value="5.1"<% If strMyODBCDriver ="5.1" Then Response.Write(" checked=""checked""") %> />
          myODBC 5.1
        </label>
        <br />
        <label>
          <input type="radio" name="myODBC" id="myODBC2" value="5.2 ANSI"<% If strMyODBCDriver ="5.2 ANSI" Then Response.Write(" checked=""checked""") %> />
         myODBC 5.2
        </label>
        <br />
        <label>
          <input type="radio" name="myODBC" id="myODBC2" value="5.3 ANSI"<% If strMyODBCDriver ="5.3 ANSI" Then Response.Write(" checked=""checked""") %> />
          myODBC 5.3
        </label>
     </td>
    </tr>
    <%
  
	End If
End If
%>
    <tr class="tableBottomRow">
      <td colspan="2"><table width="100%" border="0" cellspacing="2" cellpadding="0">
          <tr>
            <td width="42%"><input name="install" type="hidden" id="install" value="<% = Request.Form("install") %>" />
              <input name="postBack" type="hidden" id="postBack" value="True" />
              <input name="DbType" type="hidden" id="DbType" value="<% = Request.Form("DbType") %>" /></td>
            <td width="58%"><input type="button" name="testEamil" id="testEamil" value="Test Database Connection" onclick="OpenTestDbConnection(document.frmSetup)" />
              &nbsp;&nbsp;&nbsp;&nbsp;
              <input type="submit" name="button" id="button" value="Next &gt;&gt;" /></td>
          </tr>
        </table></td>
    </tr>
  </table>
</form>
<br />
<div align="center">
  <%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write("<span class=""text"" style=""font-size:10px"">Forum Software by <a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank""  style=""font-size:10px"">Web Wiz Forums&reg;</a> version " & strVersion & "</span>")
Response.Write("<br /><span class=""text"" style=""font-size:10px"">Copyright &copy;2001-2019 Web Wiz Ltd.</span>")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>
</div>
<br />
</body>
</html>
