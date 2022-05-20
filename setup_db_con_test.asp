<% @ Language=VBScript %>
<% Option Explicit %>
<!-- #include file="database/database_connection.asp" -->
<!-- #include file="functions/functions_filters.asp" -->
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




Response.Buffer = True 

'If the database is already setup then move 'em
If strDatabaseType <> "" Then Response.Redirect("default.asp")


Dim blnDetailedErrorReporting
Dim blnErrorLogging
Dim strLoggedInUsername
Dim blnLoggingEnabled
Dim strErrorMessage
Dim strAccessDbPath
Dim strAccessDbFileName
Dim strAccessOrgFileLocation
Dim blnAccessDbPhysicalPath



blnDetailedErrorReporting = True
blnErrorLogging = False
blnLoggingEnabled = False
strAccessOrgFileLocation = Server.MapPath("database\wwForum.mdb")



'Read in form contents
strSQLServerName = Trim(Mid(Request.Form("DbServer"), 1, 80))
strSQLServerName = formatSQLInput(strSQLServerName)
strSQLDBName = Trim(Mid(Request.Form("DbName"), 1, 50))
strSQLDBName = formatSQLInput(strSQLDBName)
strSQLDBUserName = Trim(Mid(Request.Form("DbUsername"), 1, 50))
strSQLDBUserName = formatSQLInput(strSQLDBUserName)
strSQLDBPassword = Trim(Mid(Request.Form("DbPassword"), 1, 50))
strSQLDBPassword = formatSQLInput(strSQLDBPassword)

strAccessDbPath = Trim(Mid(Request.Form("AccessDbPath"), 1, 80))
strAccessDbFileName = Trim(Mid(Request.Form("AccessDbName"), 1, 20))
blnAccessDbPhysicalPath = CBool(Request.Form("AccessDbServerPath"))


'Check access DB path is OK
strAccessDbPath = Replace(strAccessDbPath, "/", "\")
If strAccessDbPath <> "" AND isNull(strAccessDbPath) = False Then
	If Mid(strAccessDbPath, len(strAccessDbPath), 1) <> "\" Then  strAccessDbPath = strAccessDbPath & "\"
End If
	


'Set the type of database

'*** SQL Server ****
If Request.Form("DbType") = "SQLServer" Then
	
	strDatabaseType = "SQLServer"
	
	blnSqlSvrAdvPaging = True
	
	'Set SQL Server Driver
	If Request.Form("SQLDriver") = "SQLOLEDB" Then
		strSQLServersDriver = "SQLOLEDB"
	Else
		strSQLServersDriver = "SQL Server Native Client 11.0"
	End If
	

'*** SQL Server 2000 ****	
ElseIf Request.Form("DbType") = "SQLServer2000" Then
	
	strDatabaseType = "SQLServer"
	
	blnSqlSvrAdvPaging = False


'*** mySQL ****	
ElseIf Request.Form("DbType") = "mySQL" Then
	
	strDatabaseType = "mySQL"
	
	If Request.Form("myODBC") = "5.3 ANSI" Then
		strMyODBCDriver = "5.3 ANSI"
	ElseIf Request.Form("myODBC") = "5.2 ANSI" Then
		strMyODBCDriver = "5.2 ANSI"
	ElseIf Request.Form("myODBC") = "5.1" Then
		strMyODBCDriver = "5.1"
	Else
		strMyODBCDriver = "3.51"
	End If
	

'*** Access ****		
ElseIf Request.Form("DbType") = "Access" Then
	
	'DB Type
	strDatabaseType = "Access"
	
	If Request.Form("AccessDriver") = "JET" Then
		strAccessDriver = "Microsoft.Jet.OLEDB.4.0"
	Else
		strAccessDriver = "Microsoft.ACE.OLEDB.12.0"
	End If
	
	'Set error trapping
	On Error Resume Next
	
	'Location of database
	If blnAccessDbPhysicalPath Then
		strDbPathAndName = strAccessDbPath & strAccessDbFileName
	Else
		strDbPathAndName = Server.MapPath(strAccessDbPath & strAccessDbFileName)
	End If 
	
		
	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then	strErrorMessage = ("<br /><br /><strong>Error details:-</strong><br />" & Err.Source & "<br />" & Err.Description & "<br /><br />")
						
	'Disable error trapping
	On Error goto 0
	
	
	'Move database to new location (removed as due to support issues)
	'dim fs
	'set fs=Server.CreateObject("Scripting.FileSystemObject")
	'fs.CopyFile strAccessOrgFileLocation, Server.MapPath(strAccessDbPath & "/" & strAccessDbFileName)
	'set fs=nothing

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
	
	'Database driver
	strCon = "Provider=" & strAccessDriver & "; Data Source=" & strDbPathAndName
	
	'Database driver (Microsoft JET OLE DB driver version 4)
	'strCon = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & strDbPathAndName
		
	'Database driver (Microsoft ACE OLE DB driver) for Access 2007
	'strCon = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & strCon
		
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
If Err.Number <> 0 Then	strErrorMessage = ("<br /><br /><strong>Error details:-</strong><br />" & Err.Source & "<br />" & Err.Description & "<br /><br />")
					
'Disable error trapping
On Error goto 0

'Clean up
Set adoCon = Nothing

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" dir="ltr" lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Test Database Connection</title>
<meta name="generator" content="Web Wiz Forums" /><%

Response.Write(vbCrLf  & vbCrLf & "<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Forums(TM) " & _
vbCrLf & "Info: http://www.webwizforums.com" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->" & vbCrLf)
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******

%>
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
<link href="css_styles/default/default_style.css" rel="stylesheet" type="text/css" />
</head>
<body OnLoad="self.focus();">
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
  <tr>
    <td align="center"><h1>Test Database Connection</h1></td>
  </tr>
</table>
<table cellspacing="1" cellpadding="3" class="tableBorder" align="center" width="350">
    <tr class="tableLedger">
      <td colspan="2">Test Database Connection</td>
    <tr class="tableRow">
      <td align="left" colspan="2"><%
      	
'If error occured
If strErrorMessage <> "" Then
	
	Response.Write("<br /><strong>An Error Occurred Connecting to the Database</strong>")
	Response.Write(strErrorMessage)

Else
	Response.Write("<br /><strong>The Database Connection was Successful</strong><br />")
	
End If

%>
    <br /></td>
  </tr>
</table>
<br /><br />
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
  <tr>
    <td align="center"><input type="button" name="ok" onclick="javascript:window.close();" value="Close Window"><br />
      <br />
    </td>
  </tr>
</table>
<br />
<br />
</body>
</html>
