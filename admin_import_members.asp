<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="admin_common.asp" -->
<!--#include file="functions/functions_hash1way.asp" -->
<!--#include file="functions/functions_format_post.asp" -->
<!--#include file="functions/functions_date_time_format.asp" -->
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




'If in demo mode redirect
If blnDemoMode Then
	Call closeDatabase()
	Response.Redirect("admin_web_wiz_forums_premium.asp" & strQsSID1)
End If




'Set the script timeout to 5 hours incase there are lots of emails addresses to import
Server.ScriptTimeout = 2000000000 'secounds


'Set the response buffer to true as we maybe redirecting
Response.Buffer = False 







'Global variables
Dim lngTotalProcessed		'Counts the number of records processed




'******************************************
'***  	  Import form DB	       ****
'******************************************

'Sub procedure to read in the database subscribers
Public Sub GetDbSubscribers()
	
	Dim adoImportCon 		'Database Connection Variable
	Dim strImportCon		'Holds the connection details to db
	Dim rsImport			'Holds the imported db recordset
	Dim strDBType			'Holds the database type to import
	Dim strImpDatabaseLocation	'Holds the db location
	Dim strDatabasePassword		'Holds the db password
	Dim strDatabaseUsername		'Holds the db username
	Dim strDatabaseServer		'Holds the db server name or IP
	Dim strDatabaseName		'Holds the db database name
	Dim strDatabaseTableName	'Holds the database Table name
	Dim strDatabaseEmailField	'Holds the db email field name
	Dim strDatabaseUserNameField	'Holds the db member name field name
	Dim strDatabasePasswordField	'Holds the db password field name
	Dim strDatabasePathType		'Holds the db path type to database
	Dim lngMemberImportCount	'Counts the number of members imported
	Dim lngMemberAlreadyImported	'Counts the number of members already imported
	Dim lngNoUsername		'Counts the number of members with no email address
	Dim strEmail			'Holds the email address of the user
	Dim strUserName			'Holds the name of the user
	Dim strDbPassword		'Holds thepassword for the user
	Dim strUsersPassword		'Holds thepassword for the user
	Dim strSaltValue		'Holds the salt value
	Dim strUserCode			'Holds a user code for the user
	Dim blnMemberExists		'Set to true if the email address is already in the database
	Dim blnEmailOK			'Set to true if the email address is valid
	Dim lngMemberID			'Holds the id number of the new user
	Dim blnHTMLformat		'Holds the email format
	Dim lngTotalRecords		'Holds the total number of record to process
	Dim lngDatabaseTotalRecords
	Dim strDatabaseLocation
	Dim strDatabaseSingnature
	Dim strDatabaseNoOfPosts
	Dim strLocation
	Dim strSingnature
	Dim lngNoOfPosts
	Dim strErrorFieldName
	Dim blnUserCodeOK
	Dim intGroupID
	Dim strDatabaseFirstName
	Dim strDatabaseLastName
	Dim strFirstName
	Dim strLastName
	
	
	
	
	'Initilise variables
	lngMemberImportCount = 0
	lngMemberAlreadyImported = 0
	lngNoUsername = 0
	lngTotalProcessed = 0
	blnEmailOK = True
	blnMemberExists = false
	strUserName = ""
	strDbPassword = ""
	strUsersPassword = ""
	strSaltValue = ""
	strEmail = ""
	
	
	
	
	'Read in the form details
	strDBType = Request.Form("dbType")
	strImpDatabaseLocation = Request.Form("location")
	strDatabasePathType = Request.Form("locType")
	strDatabaseUsername = Request.Form("username")
	strDatabasePassword = Request.Form("password")
	strDatabaseServer = Request.Form("dbServerIP")
	strDatabaseName = Request.Form("dbName")
	strDatabaseTableName = Request.Form("tableName")
	strDatabaseEmailField = Request.Form("emailField")
	strDatabaseUserNameField = Request.Form("usernameField")
	strDatabasePasswordField = Request.Form("passwordField")
	strDatabaseLocation = Request.Form("where")
	strDatabaseSingnature = Request.Form("signature")
	strDatabaseNoOfPosts = Request.Form("Posts")
	intGroupID = IntC(Request.Form("GID"))
	strDatabaseFirstName = Request.Form("realNameFirst")
	strDatabaseLastName = Request.Form("realNameLast")
	
	
	
	'Create a connection odject
	Set adoImportCon = Server.CreateObject("ADODB.Connection")
	
	'If this is an access database then setup the database connection
	If strDBType = "access" OR strDBType = "access97" Then
		
		
		'If this is a path from the application to the database use the mapPath method
		If strDatabasePathType = "virtual" Then strImpDatabaseLocation = Server.MapPath(strImpDatabaseLocation)   
		
		
		'If a username and password are required then pass them across (uses slower generic db access driver
		If strDatabasePassword <> "" OR strDatabaseUsername <> "" Then
		
			strImportCon = "DRIVER={Microsoft Access Driver (*.mdb)};uid=" & strDatabaseUsername & ";pwd=" & strDatabasePassword & "; DBQ=" & strImpDatabaseLocation & "/" & strDatabaseName
		
		
		'If this is access 97 then use the jet3 db driver
		ElseIf strDBType = "access97" Then 
			
			strImportCon = "Provider=Microsoft.Jet.OLEDB.3.51; Data Source=" & strImpDatabaseLocation & "/" & strDatabaseName
		
		'Else use the jet 4 driver
		Else
			strImportCon = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & strImpDatabaseLocation & "/" & strDatabaseName
		End If
	
	
	'Else if this is MS SQL server then setup db connection string
	ElseIf strDBType = "SQLServer" Then 
	
		'MS SQL Server OLE Driver
		strImportCon = "Provider=SQLOLEDB;Server=" & strDatabaseServer & ";User ID=" & strDatabaseUsername & ";Password=" & strDatabasePassword & ";Database=" & strDatabaseName & ";"
	
	'Else if this is mySQL then setup db connection string
	ElseIf strDBType = "mySQL" Then 
	
		'My SQL ODBC Driver
		strImportCon = "Driver={MySQL ODBC 3.51 Driver};Server=" & strDatabaseServer & ";Port=3306;Option=4;Database=" & strDatabaseName & ";Uid=" & strDatabaseUsername & ";Pwd=" & strDatabasePassword & ";"
	End If
	
	
	
	'Set error trapping
	On Error Resume Next
	
	'Open database connection
	adoImportCon.connectionstring = strImportCon
	
	'Set an active connection to the Connection object
	adoImportCon.Open
	
	'If an error has occurred while connecting to database let the user know
	If Err.Number <> 0 Then
	
		Response.Write("<script language=""JavaScript"">" & _
		vbCrLf & "	document.getElementById('displayState').innerHTML = 'Database import process stopped. See detailed error message below.';" & _
		vbCrLf & "	document.getElementById('errMsg').value = document.getElementById('errMsg').value +  'Error connecting to database,\nError: " & errorDescription(err.description) & "\n';" & _
		vbCrLf & "</script>")
		
		Call closeDatabase()
		
		Response.Flush
		Response.End
	End If
	
	

	
	'Get details from database
	Set rsImport = Server.CreateObject("ADODB.Recordset")
	
	
	
	'First count the number of subscribers to import
	strSQL = "SELECT COUNT(*) AS TotalRecords FROM " & strDatabaseTableName & ";"
	
	
	'Query the database
	rsImport.Open strSQL, adoImportCon
	
	'If an error has occurred while getting table data let the user know
	If Err.Number <> 0 Then
	
		Response.Write("<script language=""JavaScript"">" & _
		vbCrLf & "	document.getElementById('displayState').innerHTML = 'Database import process stopped. See detailed error message below.';" & _
		vbCrLf & "	document.getElementById('errMsg').value = document.getElementById('errMsg').value +  'Error, incorrect table name,\nError: " & errorDescription(err.description) & "\n';" & _
		vbCrLf & "</script>")
		
		Call closeDatabase()
		
		Response.Flush
		Response.End
	End If
	
	'Disable error trapping
	On Error goto 0
	
	
	'Get the totla records from db
	lngTotalRecords = CLng(rsImport("TotalRecords"))
	
	'Display on page number of subscribers to import
	Response.Write("<script language=""JavaScript"">" & _
	vbCrLf & "	document.getElementById('displayState').innerHTML = 'Initialising database import process...';" & _
	vbCrLf & "	document.getElementById('subscribers').innerHTML = '" & lngTotalRecords & "';" & _
	vbCrLf & "</script>")
	
	'Close the recordset
	rsImport.Close
	
	
	
	'Build SQL query
	strSQL = "SELECT * FROM " & strDatabaseTableName & ";"
	
	'Query the database
	rsImport.Open strSQL, adoImportCon
	
	
	'Loop through recordset
	Do While NOT  rsImport.EOF
	
		'Initilise variables
		blnEmailOK = True
		blnMemberExists = false
		strErrorFieldName = ""
		blnUserCodeOK = false
		
		'Count the number of records processed
		lngTotalProcessed = lngTotalProcessed + 1
		
		
		'Set error trapping
		On Error Resume Next
	
		'Read in the details from the database
	
		strUserName = rsImport(strDatabaseUserNameField)
		If strUserName <> "" Then strUserName = formatSQLInput(strUserName)
		If Err.Number <> 0 Then strErrorFieldName = strErrorFieldName & "\'Name Source Field\', "
		Err.Number = 0
		
		If strDatabasePasswordField <> "" Then 
			strDbPassword = rsImport(strDatabasePasswordField)
			If strDbPassword <> "" Then strDbPassword = removeAllTags(strDbPassword)
			If Err.Number <> 0 Then strErrorFieldName = strErrorFieldName & "\'Password Source Field\', "
		End If
		Err.Number = 0
		
		If strDatabaseEmailField <> "" Then 
			strEmail = LCase(rsImport(strDatabaseEmailField))
			If strEmail <> "" Then strEmail = removeAllTags(strEmail)
			If Err.Number <> 0 Then strErrorFieldName = "\'Email Address Source Field\', "
		End If
		Err.Number = 0
		
		If strDatabaseLocation <> "" Then 
			strLocation = rsImport(strDatabaseLocation)
			If strLocation <> "" Then strLocation = removeAllTags(strLocation)
			If Err.Number <> 0 Then strErrorFieldName = strErrorFieldName & "\'Location Source Field\', "
		End If
		Err.Number = 0
		
		If strDatabaseNoOfPosts <> "" Then 
			lngNoOfPosts = rsImport(strDatabaseNoOfPosts)
			If lngNoOfPosts <> "" Then lngNoOfPosts = CLng(lngNoOfPosts)
			If Err.Number <> 0 Then strErrorFieldName = strErrorFieldName & "\'No Of Posts Source Field\', "
		End If
		Err.Number = 0
		
		If strDatabaseSingnature <> "" Then 
			strSingnature = rsImport(strDatabaseSingnature)
			If strSingnature <> "" Then strSingnature = HTMLsafe(strSingnature)
			If Err.Number <> 0 Then strErrorFieldName = strErrorFieldName & "\'Signature Source Field\', "
		End If
		Err.Number = 0
		
		If strDatabaseFirstName <> "" Then 
			strFirstName = rsImport(strDatabaseFirstName)
			If strFirstName <> "" Then strFirstName = HTMLsafe(strFirstName)
			If Err.Number <> 0 Then strErrorFieldName = strErrorFieldName & "\'First Name Source Field\', "
		End If
		Err.Number = 0
		
		If strDatabaseLastName <> "" Then 
			strLastName = rsImport(strDatabaseLastName)
			If strLastName <> "" Then strLastName = HTMLsafe(strLastName)
			If Err.Number <> 0 Then strErrorFieldName = strErrorFieldName & "\'Last Name Source Field\', "
		End If
		Err.Number = 0
		
		
		
		
		'If an error has occurred while getting data let the user know
		If strErrorFieldName <> "" Then
		
			Response.Write("<script language=""JavaScript"">" & _
			vbCrLf & "	document.getElementById('displayState').innerHTML = 'Database import process stopped. See detailed error message below.';" & _
			vbCrLf & "	document.getElementById('errMsg').value = document.getElementById('errMsg').value +  'Error, incorrect " & strErrorFieldName & "\nError: " & errorDescription(err.description) & "\n';" & _
			vbCrLf & "</script>")
			
			Call closeDatabase()
			
			Response.Flush
			Response.End
		End If
		
		'Disable error trapping
		On Error goto 0
		
		
		
		
		'If no email address then increament the no email address count
		If strUserName = "" OR isNull(strUserName) Then
			
			lngNoUsername = lngNoUsername + 1
		
		
	
		'Run if email address is returned	
		Else
			
			
			
			'Initalise the strSQL variable with an SQL statement to query the database
			strSQL = "SELECT " & strDbTable & "Author.* " & _
				"FROM " & strDbTable & "Author " & _
				"WHERE " & strDbTable & "Author.Username = '" & strUserName & "';"
				
			
			'Remove SQL safe single quote double up set in the format SQL function
                        strUsername = Replace(strUsername, "''", "'", 1, -1, 1)
                        strUsername = Replace(strUsername, "\'", "'", 1, -1, 1)
			
			With rsCommon
			
				'Set the cursor type property of the record set to Forward Only
				.CursorType = 0
						
				'Set the Lock Type for the records so that the record set is only locked when it is updated
				.LockType = 3
						
				'Query the database
				.Open strSQL, adoCon
				
				'If a record is returned this email address is already in the database
				If NOT .EOF Then 
					blnMemberExists = true
					
					'Increment the already imported number
					lngMemberAlreadyImported = lngMemberAlreadyImported + 1
				End If

				
				
				'If the member doesn't already exist then enter them into the db
				If blnMemberExists = False Then
					
					
					'Create password if there are none
					If strDbPassword = "" Then 
						strUsersPassword = hexValue(7)
					Else
						strUsersPassword = strDbPassword
					End If
						
					'If the passowrds need to be encrypted then create a slat value and encrypt passords
					If blnEncryptedPasswords Then
						
						'generate a salt value
						strSaltValue = hexValue(8)
						
						'Concatenate salt value to the password
				                strUsersPassword = strUsersPassword & strSaltValue
								
						'Encrypt the  password
						strUsersPassword = HashEncode(strUsersPassword)
					End If
					
					
						
					'Add new record to a new recorset
					.AddNew
					
					'Set database fields
					.Fields("Username") = Trim(Mid(strUserName, 1, 20))
					.Fields("Password") = strUsersPassword
					If blnEncryptedPasswords Then .Fields("Salt") = strSaltValue
					.Fields("User_code") = userCode(strUsername)
					
					.Fields("Author_email") = Trim(Mid(strEmail, 1, 50))
					.Fields("Group_ID") = intGroupID
					.Fields("Join_date") = internationalDateTime(Now())
					.Fields("Last_visit") = internationalDateTime(Now())
					.Fields("Banned") = False
					.Fields("Info") = "" 'This is to prevent errors in mySQL
					.Fields("Active") = True
					
					If strDatabaseLocation <> "" Then .Fields("Location") = Trim(Mid(strLocation, 1, 60))
					If strDatabaseSingnature <> "" Then .Fields("Signature") = Trim(Mid(strSingnature, 1, 245))
					If strDatabaseNoOfPosts <> "" Then .Fields("No_of_posts") = CLng(lngNoOfPosts)
					If strDatabaseFirstName <> "" Then .Fields("Real_name") = Trim(Mid(strFirstName & " " & strLastName, 1, 30))
						
					
					
					.Fields("Date_format") = saryDateTimeData(1,0)
		                        .Fields("Time_offset") = saryDateTimeData(19,0)
		                        .Fields("Time_offset_hours") = saryDateTimeData(20,0)
		                        .Fields("Reply_notify") = False
		                        .Fields("Rich_editor") = blnRTEEditor
		                        .Fields("PM_notify") = False
		                        .Fields("Show_email") = False
		                        .Fields("Attach_signature") = True
					
					
						
					'Update the database
					.Update	
					
					
					'Increment the number of users imported by 1
					lngMemberImportCount = lngMemberImportCount + 1
					
				End If
				
				'Close rs
				.Close
				
			End With
		End If
		
		'Display on page number of subscribers to import
		Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
		vbCrLf & "	document.getElementById('displayState').innerHTML = 'Importing member...';" & _			
		vbCrLf & "	document.getElementById('imported').innerHTML = '" & lngMemberImportCount & "';" & _
		vbCrLf & "	document.getElementById('done').innerHTML = '" & lngMemberAlreadyImported & "';" & _
		vbCrLf & "	document.getElementById('noname').innerHTML = '" & lngNoUsername & "';" & _
		vbCrLf & "	document.getElementById('total').innerHTML = '" & lngTotalProcessed & "';" & _
		vbCrLf & "	document.getElementById('progress').innerHTML = '" & percentageCalculate(lngTotalProcessed, lngTotalRecords, 0) & "';" & _
		vbCrLf & "	document.getElementById('progressBar').style.width = '" & percentageCalculate(lngTotalProcessed, lngTotalRecords, 3) & "';" & _
		vbCrLf & "</script>")
					
		
	
		'Move to next record
		rsImport.MoveNext
	Loop
	
	
	'Display on page number of subscribers to import
	Response.Write("<script language=""JavaScript"">" & _
	vbCrLf & "	document.getElementById('displayState').innerHTML = 'Database import process complete.';" & _
	vbCrLf & "</script>")
	
	
	'Clean up
	adoImportCon.Close
	Set adoImportCon = Nothing

End Sub




'******************************************
'***  	    Calculate Percentage       ****
'******************************************
Private Function percentageCalculate(ByRef lngNumberProcessed, ByRef lngTotalToProcess, ByRef intDecPlaces)
	
	'If there are no newsletters sent yet then format the percent by 0 otherwise an overflow error will happen
	If lngTotalProcessed = 0 Then
		percentageCalculate = FormatPercent(0, 0)
	
	'Else read in the the percentage of newsletters sent
	Else
		percentageCalculate = FormatPercent((lngNumberProcessed / lngTotalToProcess), intDecPlaces)
	End If
End Function




'******************************************
'***  	    Format Error Description   ****
'******************************************
Private Function errorDescription(strErrorDescription)

	'Format the error description for javascrip
 	strErrorDescription = Replace(strErrorDescription, vbCrLf, "", 1, -1, 1)
 	strErrorDescription = Replace(strErrorDescription, "\", "\\", 1, -1, 1)
 	strErrorDescription = Replace(strErrorDescription, "'", "\'", 1, -1, 1)
 	
 	'Return the function result
 	errorDescription = strErrorDescription
End Function





	
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Import Members</title>
<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write("<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Forums(TM) ver. " & strVersion & "" & _
vbCrLf & "Info: http://www.webwizforums.com" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
</head>
<body>
<!-- #include file="includes/admin_header_inc.asp" -->
   <h1>Import Members </h1>
   <a href="admin_menu.asp" target="_self">Return to the the Admin Control Panel Menu</a><br />
   <br />
   Youmembers  are being Imported.<br />
   <span class="lgText">Do not close this window while this task is being carried out. </span><br />
   <br />
   <table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="tableBorder">
    <tr>
     <td align="left" class="tableLedger">Importing members... </td>
    </tr>
    <tr class="tableRow">
     <td align="left"><table width="72%" border="0" align="center" cellpadding="6" cellspacing="0">
       <tr>
        <td width="20%" rowspan="10" class="tableRow">&nbsp;</td>
        <td width="7%" class="tableRow"><span id="subscribers"></span></td>
        <td width="53%" class="tableRow">Total Records Found</td>
        <td width="20%" rowspan="8" class="tableRow">&nbsp;</td>
       </tr>
       <tr>
        <td class="tableRow"><span id="noname"></span></td>
        <td class="tableRow">Records With No Username</td>
       </tr>
       <tr>
        <td class="tableRow"><span id="done"></span></td>
        <td class="tableRow">Members Already Imported</td>
       </tr>
       <tr>
        <td class="tableRow"><span id="imported"></span></td>
        <td class="tableRow">Members Imported </td>
       </tr>
       <tr>
        <td class="tableRow"><strong><span id="total"></span></strong></td>
        <td class="tableRow"><strong>Total Processed </strong></td>
       </tr>
       <tr>
        <td class="tableRow">&nbsp;</td>
        <td class="tableRow">&nbsp;</td>
       </tr>
       <tr>
        <td colspan="2" class="tableRow"><span id="progress">0%</span> Progress</td>
       </tr>
       <tr>
        <td colspan="2" class="tableRow"><table width="300" border="0" cellpadding="0" cellspacing="1" bgcolor="#999999">
          <tr>
           <td height="17" background="<% = strImagePath %>progress_bar_bg.gif"><img src="<% = strImagePath %>progress_bar.gif" alt="Progress Bar" height="17" style="width:0%;" id="progressBar" /></td>
          </tr>
         </table></td>
       </tr>
       <tr>
        <td colspan="3" class="tableRow"><strong>Status: <span id="displayState"></span></strong></td>
       </tr>
      </table></td>
    </tr>
    <tr>
     <td align="left" class="tableLedger">Error Details</td>
    </tr>
    <tr>
     <td align="left" class="tableRow">Below are the details of any error messages returned by the server<br />
      <textarea name="errMsg" cols="80" rows="10" id="errMsg" readonly="readonly">Error messages:-

</textarea></td>
    </tr>
   </table>
   <strong><br />
   <br />
   <br />
   <!-- #include file="includes/admin_footer_inc.asp" -->
<%


'Call sub to do databse import
Call GetDbSubscribers()


'Clean up
Call closeDatabase()
 

%>
