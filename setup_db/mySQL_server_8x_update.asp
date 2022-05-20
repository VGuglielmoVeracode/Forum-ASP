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




'No Warranty
'----------- 
'There is no warranty for this program, to the extent permitted by applicable law, except when otherwise stated in writing the copyright holders and/or other parties provide the program ‘AS IS’ without warranty of any kind, either expressed or implied, including, but not limited to, the implied warranties of merchantability and fitness for a particular purpose. The entire risk as to the quality and performance of the program is with you. Should the program prove defective, you assume the cost of all necessary servicing, repair or correction. 
'
'In no event unless required by applicable law or agreed to in writing will any copyright holder, be liable to you for damages, including any general, special, incidental or consequential damages arising out of the use or inability to use the program (including but not limited to loss of data or data being rendered inaccurate or losses sustained by you or third parties or a failure of the program to operate with any other programs), even if such holder or other party has been advised of the possibility of such damages.
'




'Check file is being called from setup_db-asp
If NOT blnSetupDB = True Then
	
	Response.Write("Access to this file is not permitted")
	
	'Stop processing script
	Response.Flush
	Response.End
	
End If




'Let the user know the database is being created
Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
vbCrLf & "	document.getElementById('displayState').innerHTML = 'Your Database is being updated. Please be patient as this may take a few minutes to complete.';" & _
vbCrLf & "</script>")






'Resume on all errors
On Error Resume Next

'intialise variables
blnErrorOccured = False


'If a username and password is entred then start the ball rolling
If strDatabaseType = "mySQL" AND strSQLDBUserName <> "" Then
	
	

	'Open the database
	Call openDatabase(strCon)

	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then
		
		
		
		Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
		vbCrLf & "	document.getElementById('displayState').innerHTML = '<img src=""forum_images/error.png"" alt=""Error"" /> <strong>Error</strong><br /><strong>Error Connecting to database on mySQL Server</strong><br /><br />Replace the database/database_settings.asp file with the one from the orginal Web Wiz Forums download and start the setup process again.<br /><br /><strong>Error Details:</strong><br />" & Err.description & "';" & _
		vbCrLf & "</script>")

		
	Else
		
		'Intialise the main ADO recordset object
		Set rsCommon = CreateObject("ADODB.Recordset")
		
		
		'Check to see if the web wiz forums database exists
		
		'Get the admin account
		strSQL = "SELECT " & strDbTable & "Author.Username " & _
		"FROM " & strDbTable & "Author " & _
		"WHERE " & strDbTable & "Author.Author_ID = 1;"
		
		'Query the database
		rsCommon.Open strSQL, adoCon
		
		'If error occured the database has not been created
		If NOT CLng(Err.Number) = 0 Then
			
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "	document.getElementById('displayState').innerHTML = '<img src=""forum_images/error.png"" alt=""Error"" /> <strong>Error</strong><br /><strong>The Database Setup Wizard has can not find a previous Web Wiz Forums tables in the database to update.</strong><br /><br />Replace the database/database_settings.asp file with the one from the orginal Web Wiz Forums download and start the setup process again.';" & _
			vbCrLf & "</script>")
			
			
			Set rsCommon = Nothing
			Set adoCon = Nothing
			Response.End
		
		
		End If
		
		'Reset error object
		rsCommon.Close
		Err.Number = 0
		
		
		
		
	
	
		'Check to see if the database is already updated
		
		'Get the admin account
		strSQL = "SELECT " & strDbTable & "Author.Gender, " & strDbTable & "Author.Photo " & _
		"FROM " & strDbTable & "Author " & _
		"WHERE " & strDbTable & "Author.Author_ID = 1;"
		
		'Query the database
		rsCommon.Open strSQL, adoCon
		
		'If no error occured the database has been created
		If CLng(Err.Number) = 0 Then
			
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "	document.getElementById('displayState').innerHTML = '<img src=""forum_images/error.png"" alt=""Error"" /> <strong>Error</strong><br /><strong>The Database Setup Wizard has detected that your database has already been updated.</strong><br /><br />Click here to go to your <a href=""default.asp"">Web Wiz Forum Homepage</a>.';" & _
			vbCrLf & "</script>")
			
			
			Set rsCommon = Nothing
		
		'Create the database
		Else

			'Reset error object
			Err.Number = 0
			Set rsCommon = Nothing
	
	
			
		
		
			
		
	
	
	'******************************************
	'***  	Update/Create the tables      *****
	'******************************************
	
			'Stage one start
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br /><strong>Stage 1: Creating New Database Tables and Fields..... </strong>';" & _
			vbCrLf & "</script>")
			
			
			
			'Update tblAuthor
			strSQL = "ALTER TABLE " & strDbTable & "Author  "
			strSQL = strSQL & "ADD Gender varchar(10) NULL, "
			strSQL = strSQL & "ADD Photo varchar(100) NULL, "
			strSQL = strSQL & "ADD Newsletter tinyint(1) NOT NULL DEFAULT '-1';"
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br />Error Altering the Table " & strDbTable & "Group <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			'Update tblForum
			strSQL = "ALTER TABLE " & strDbTable & "Forum "
			strSQL = strSQL & "ADD Last_topic_ID INT NOT NULL DEFAULT '0'; "
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br />Error Altering the Table " & strDbTable & "Forum <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			'Update tblTopic
			strSQL = "ALTER TABLE " & strDbTable & "Topic "
			strSQL = strSQL & "ADD Event_date_end datetime NULL; "
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br />Error Altering the Table " & strDbTable & "Topic <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			'Update tblGroup
			strSQL = "ALTER TABLE " & strDbTable & "Group "
			strSQL = strSQL & "ADD Image_uploads tinyint(1) NOT NULL DEFAULT '0', "
			strSQL = strSQL & "ADD File_uploads tinyint(1) NOT NULL DEFAULT '0'; "
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br />Error Altering the Table " & strDbTable & "Group <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			'Update tblConfiguration
			strSQL = "ALTER TABLE " & strDbTable & "Configuration "
			strSQL = strSQL & "ADD A_code tinyint(1) NOT NULL DEFAULT '0', "
			strSQL = strSQL & "ADD Upload_allocation smallint NULL, "
			strSQL = strSQL & "ADD NewsPad tinyint(1) NOT NULL DEFAULT '0', "
			strSQL = strSQL & "ADD NewsPad_URL varchar(50) NULL; "
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br />Error Altering the Table " & strDbTable & "Configuration <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
	
	
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<strong>Complete</strong>';" & _
			vbCrLf & "</script>")
			
			
			
	
	'******************************************
	'***  	Insert default values	      *****
	'******************************************
			
			
			'Stage 2 start
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br /><strong>Stage 2: Entering default values for new fields..... </strong>';" & _
			vbCrLf & "</script>")
	
			
			'Enter the default values in the Group Table
			strSQL = "UPDATE " & strDbTable & "Group " & _
			"SET " & _
			strDbTable & "Group.Image_uploads = " & strDBFalse & ", " & _
			strDbTable & "Group.File_uploads = " & strDBFalse & "; "
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br />Error entering default values in the Table " & strDbTable & "Configuration<br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
	
	
			
			'Enter the default values in the Configuration Table
			strSQL = "UPDATE " & strDbTable & "Configuration " & _
			"SET " & _
			strDbTable & "Configuration.Upload_allocation = 10, " & _
			strDbTable & "Configuration.Title_image = 'forum_images/web_wiz_forums.png', " & _
			strDbTable & "Configuration.Skin_file = 'css_styles/default/', " & _
			strDbTable & "Configuration.Skin_image_path = 'forum_images/', " & _
			strDbTable & "Configuration.Skin_nav_spacer = ' &gt; ', " & _
			strDbTable & "Configuration.Flash = " & strDBTrue & ", " & _
			strDbTable & "Configuration.A_code = " & strDBTrue & ", " & _
			strDbTable & "Configuration.L_Code = " & strDBTrue & ", " & _
			strDbTable & "Configuration.NewsPad_URL = '', " & _
			strDbTable & "Configuration.NewsPad = " & strDBTrue & "; "
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br />Error entering default values in the Table " & strDbTable & "Configuration<br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<strong>Complete</strong>';" & _
			vbCrLf & "</script>")
			
			
			
			
			
	
			'Display a message to say the database is created
			If blnErrorOccured = True Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />" & Err.description & "<br /><br /><h2>mySQL database is updated, but with Error!</h2>'" & _
				vbCrLf & "</script>")
			
			End If
			
			
		
		End If
	End If

	'Reset Server Variables
	Set adoCon = Nothing
End If
%>
      