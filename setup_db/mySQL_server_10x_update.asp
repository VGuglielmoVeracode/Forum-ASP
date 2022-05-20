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




'Check file is being called from setup_db-asp
If NOT blnSetupDB = True Then
	
	Response.Write("Access to this file is not permitted")
	
	'Stop processing script
	Response.Flush
	Response.End
	
End If





'Let the user know the database is being created
If Request.QueryString("setup") = "SqlServer10Update" Then
	Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
	vbCrLf & "	document.getElementById('displayState').innerHTML = 'Your Database is being updated. Please be patient as this may take a few minutes to complete.';" & _
	vbCrLf & "</script>")

End If









'Resume on all errors
On Error Resume Next


'intialise variables
blnErrorOccured = False


	

'Open the database
Call openDatabase(strCon)

'If an error has occurred write an error to the page
If Err.Number <> 0 Then
		
		
		
	Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
	vbCrLf & "	document.getElementById('displayState').innerHTML = '<img src=""forum_images/error.png"" alt=""Error"" /> <strong>Error</strong><br /><strong>Error Connecting to database</strong><br /><br />Replace the database/database_settings.asp file with the one from the orginal Web Wiz Forums download and start the setup process again.<br /><br /><strong>Error Details:</strong><br />" & Err.description & "';" & _
	vbCrLf & "</script>")

		
Else
		
		
	'Check to see if the database is already created

	'Intialise the main ADO recordset object
	Set rsCommon = CreateObject("ADODB.Recordset")
	
	'Get the admin account
	strSQL = "SELECT " & strDbTable & "Author.Username " & _
	"FROM " & strDbTable & "Author " & _
	"WHERE " & strDbTable & "Author.Author_ID = 1;"

	'Query the database
	rsCommon.Open strSQL, adoCon
	
	'If error occured the database has been created
	If NOT CLng(Err.Number) = 0 Then
		
		Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
		vbCrLf & "	document.getElementById('displayState').innerHTML = '<img src=""forum_images/error.png"" alt=""Error"" /> <strong>Error</strong><br /><strong>The Database Setup Wizard can not find any Web Wiz Forums tables in the database.</strong><br /><br />Replace the database/database_settings.asp file with the one from the orginal Web Wiz Forums download and start the setup process again.';" & _
		vbCrLf & "</script>")
		
		
		Set rsCommon = Nothing
	
	'Create the database
	Else

		
		
		'Altering Thread
		strSQL = "ALTER TABLE tblThread MODIFY  IP_addr varchar(50)"
			
		'Write to the database
		adoCon.Execute(strSQL)
	
		'If an error has occured write an error to the page
		If Err.Number <> 0 Then
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br />Error Altering the Table " & strDbTable & "Thread <br />" & Err.description & ".';" & _
			vbCrLf & "</script>")
		
			'Reset error object
			Err.Number = 0
	
			'Set the error boolean to True
			blnErrorOccured = True
		End If
		
		
		
		
		
		
		'Reset error object
		Set rsCommon = Nothing



		Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
		vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<strong>Complete</strong>';" & _
		vbCrLf & "</script>")
			
			
			
			
	
		'Display a message to say the database is created
		If blnErrorOccured = True Then
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />" & Err.description & "<br /><br /><h2>SQL Server database is updated, but with Error!</h2>'" & _
			vbCrLf & "</script>")
		Else
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br /><h2>Congratulations, Web Wiz Forums Database update is now complete</h2>'" & _
			vbCrLf & "</script>")
		End If
			
			
			
		'Display completed message
		Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
		vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Click here to go to your <a href=""default.asp"">Forum Homepage</a><br />Click here to login to your <a href=""admin.asp"">Forum Admin Area</a>'" & _
     		vbCrLf & "</script>")
	
	End If
End If

'Reset Server Variables
Set adoCon = Nothing

%>