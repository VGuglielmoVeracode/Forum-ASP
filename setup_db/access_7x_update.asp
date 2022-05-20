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
If strDatabaseType = "Access" Then
	
	
	'Open the database
	Call openDatabase(strCon)

	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then
		
		
		
		Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
		vbCrLf & "	document.getElementById('displayState').innerHTML = '<img src=""forum_images/error.png"" alt=""Error"" /> <strong>Error</strong><br /><strong>Error Connecting to Access database</strong><br /><br />Click back on your web browser and check that you have entered the correct path to the Web Wiz Forums Access Database.<br /><br /><strong>Error Details:</strong><br />" & Err.description & "';" & _
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
		strSQL = "SELECT " & strDbTable & "Author.Info " & _
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
	
	
			'Update tblForum
			strSQL = "ALTER TABLE [" & strDbTable & "Forum] ADD "
			strSQL = strSQL & "[Sub_ID] INTEGER NOT NULL DEFAULT 0, "
			strSQL = strSQL & "[Last_post_author_ID] INTEGER NOT NULL DEFAULT 1, "
			strSQL = strSQL & "[Last_post_date] DATETIME NULL DEFAULT Now(), "
			strSQL = strSQL & "[Last_topic_ID] INTEGER DEFAULT 0 "
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
	
				'Write an error message
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Updating the Table " & strDbTable & "Forum <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
	
	
	
			'Update tblTopic
			strSQL = "ALTER TABLE [" & strDbTable & "Topic] ADD "
			strSQL = strSQL & "[Icon] VARCHAR (20),"
			strSQL = strSQL & "[Start_Thread_ID] INTEGER NULL, "
			strSQL = strSQL & "[Last_Thread_ID] INTEGER NULL ,"
			strSQL = strSQL & "[No_of_replies] INTEGER NOT NULL DEFAULT 0, "
			strSQL = strSQL & "[Hide] YESNO NOT NULL DEFAULT 0,"
			strSQL = strSQL & "[Event_date] DATETIME NULL, "
			strSQL = strSQL & "[Event_date_end] DATETIME NULL "
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
	
				'Write an error message
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Updating the Table " & strDbTable & "Topic <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
	
	
			
			'Update tblThread
			strSQL = "ALTER TABLE [" & strDbTable & "Thread] ADD "
			strSQL = strSQL & "[Hide] YESNO NOT NULL DEFAULT FALSE "
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Update the Table " & strDbTable & "Thread <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			'Update tblPermissions
			strSQL = "ALTER TABLE [" & strDbTable & "Permissions] ADD "
			strSQL = strSQL & "[View_Forum] YESNO NOT NULL DEFAULT TRUE, "
			strSQL = strSQL & "[Display_post] YESNO NOT NULL DEFAULT FALSE, "
			strSQL = strSQL & "[Calendar_event] YESNO NOT NULL DEFAULT FALSE"
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Updating the Table " & strDbTable & "Permissions <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			'Update tblAuthor
			strSQL = "ALTER TABLE [" & strDbTable & "Author] ADD "
			strSQL = strSQL & "[Gender] VARCHAR (10) NULL ,"
			strSQL = strSQL & "[Photo] VARCHAR (100) NULL ,"
			strSQL = strSQL & "[No_of_PM] INTEGER NOT NULL DEFAULT 0,"
			strSQL = strSQL & "[Skype] VARCHAR (30) NULL ,"
			strSQL = strSQL & "[Login_attempt] INTEGER NOT NULL DEFAULT 0, "
			strSQL = strSQL & "[Banned] YESNO NOT NULL DEFAULT FALSE, "
			strSQL = strSQL & "[Info] VARCHAR (255) NULL, "
			strSQL = strSQL & "[Newsletter] YESNO NOT NULL DEFAULT 1 "
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Altering the Table " & strDbTable & "Author <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			'Update tblConfiguration
			strSQL = "ALTER TABLE [" & strDbTable & "Configuration] ADD "
			strSQL = strSQL & "[ID] AUTOINCREMENT, "
			strSQL = strSQL & "[Mail_username] VARCHAR (50) NULL, "
			strSQL = strSQL & "[Mail_password] VARCHAR (50) NULL, "
			strSQL = strSQL & "[Topic_icon] YESNO NOT NULL DEFAULT TRUE, "
			strSQL = strSQL & "[Long_reg] YESNO NOT NULL DEFAULT TRUE, "
			strSQL = strSQL & "[CAPTCHA] YESNO NOT NULL DEFAULT TRUE , "
			strSQL = strSQL & "[Skin_file] VARCHAR (50) NULL, "
			strSQL = strSQL & "[Skin_image_path] VARCHAR (50) NULL, "
			strSQL = strSQL & "[Skin_nav_spacer] VARCHAR (15) NULL, "
			strSQL = strSQL & "[Guest_SID] YESNO NOT NULL DEFAULT TRUE , "
			strSQL = strSQL & "[Calendar] YESNO NOT NULL DEFAULT TRUE, "
			strSQL = strSQL & "[Member_approve] YESNO NOT NULL DEFAULT FALSE, "
			strSQL = strSQL & "[RSS] YESNO NOT NULL DEFAULT TRUE, "
			strSQL = strSQL & "[Install_ID] VARCHAR (15) NULL, "
			strSQL = strSQL & "[PM_Flood] INTEGER NOT NULL DEFAULT 10, "
			strSQL = strSQL & "[A_code] YESNO NOT NULL DEFAULT 0, "
			strSQL = strSQL & "[Upload_allocation] INTEGER NULL, "
			strSQL = strSQL & "[NewsPad] YESNO NOT NULL DEFAULT 0, "
			strSQL = strSQL & "[NewsPad_URL] VARCHAR (50) NULL "
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Altering the Table " & strDbTable & "Configuration <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			'Update tblDateTimeFormat
			strSQL = "ALTER TABLE [" & strDbTable & "DateTimeFormat] ADD "
			strSQL = strSQL & "[ID] AUTOINCREMENT, "
			strSQL = strSQL & "[Time_offset] VARCHAR (1) NOT NULL ,"
			strSQL = strSQL & "[Time_offset_hours] INTEGER NOT NULL DEFAULT 0"
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Altering the Table " & strDbTable & "DateTimeFormat <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			'Update tblGroup
			strSQL = "ALTER TABLE [" & strDbTable & "Group] ADD "
			strSQL = strSQL & "[Image_uploads] YESNO NOT NULL DEFAULT 0, "
			strSQL = strSQL & "[File_uploads] YESNO NOT NULL DEFAULT 0 "
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Altering the Table " & strDbTable & "Group <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			
			'Create the Poll Vote Table
			strSQL = "CREATE TABLE [" & strDbTable & "PollVote] ("
			strSQL = strSQL & "[Poll_ID] INTEGER NOT NULL DEFAULT 0,"
			strSQL = strSQL & "[Author_ID] INTEGER NOT NULL DEFAULT 0"
			strSQL = strSQL & ")"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "PollVote <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
	
	
			
			'Create the Thread Table
			strSQL = "CREATE TABLE [" & strDbTable & "Session] ("
			strSQL = strSQL & "[Session_ID] VARCHAR (50) PRIMARY KEY NOT NULL,"
			strSQL = strSQL & "[IP_address] VARCHAR (50) NOT NULL, "
			strSQL = strSQL & "[Last_active] DATETIME NOT NULL DEFAULT Now() ,"
			strSQL = strSQL & "[Session_data] VARCHAR (255) NULL "
			strSQL = strSQL & ")"
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "Session <br />" & Err.description & ".';" & _
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
	
			
	
			'Enter the default values in the Forum Table
			strSQL = "UPDATE " & strDbTable & "Forum " & _
			"SET " & strDbTable & "Forum.Sub_ID = 0, " & _
			strDbTable & "Forum.Show_topics = 0 "
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error entering default values in the Table " & strDbTable & "Forum<br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
	
			'Enter the default values in the Author Table
			strSQL = "UPDATE " & strDbTable & "Author " & _
			"SET " & _
			strDbTable & "Author.No_of_PM = 0, " & _
			strDbTable & "Author.Login_attempt = 0, " & _
			strDbTable & "Author.Banned = False;"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error entering default values in the Table " & strDbTable & "Author<br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			'Enter the default values in the DateTimeFormat Table
			strSQL = "UPDATE " & strDbTable & "DateTimeFormat " & _
			"SET " & _
			strDbTable & "DateTimeFormat.Time_offset = '+', " & _
			strDbTable & "DateTimeFormat.Time_offset_hours = 0;"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error entering default values in the Table " & strDbTable & "DateTimeFormat<br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
	
			
			'Enter the default values in the Configuration Table
			strSQL = "UPDATE " & strDbTable & "Configuration " & _
			"SET " & _
			strDbTable & "Configuration.Text_link = False, " & _
			strDbTable & "Configuration.Topic_icon = True, " & _
			strDbTable & "Configuration.Long_reg = True, " & _
			strDbTable & "Configuration.CAPTCHA = True, " & _
			strDbTable & "Configuration.Guest_SID = False, " & _
			strDbTable & "Configuration.Calendar = True, " & _
			strDbTable & "Configuration.RSS = True, " & _
			strDbTable & "Configuration.PM_Flood = 10, " & _
			strDbTable & "Configuration.Upload_allocation = 10, " & _
			strDbTable & "Configuration.Title_image = 'forum_images/web_wiz_forums.png', " & _
			strDbTable & "Configuration.Skin_file = 'css_styles/default/', " & _
			strDbTable & "Configuration.Skin_image_path = 'forum_images/', " & _
			strDbTable & "Configuration.Skin_nav_spacer = ' &raquo; ', " & _
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
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error entering default values in the Table " & strDbTable & "Configuration<br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
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
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error entering default values in the Table " & strDbTable & "Configuration<br />" & Err.description & ".';" & _
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
	'***  	Populate Updated Database     *****
	'******************************************
	
			'Stage 3 start
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br /><strong>Stage 3: Populating Database with Default Forum Permissions (this may take some time to complete).... </strong>';" & _
			vbCrLf & "</script>")
			
	
			
			'Intialise the main ADO recordset object
			Set rsCommon = Server.CreateObject("ADODB.Recordset")
			
			
			'Get the forum ID's from the database as these are used for mutiple procedures
			strSQL = "SELECT " & strDbTable & "Forum.Forum_ID FROM " & strDbTable & "Forum;"
			
			'Get rs
			rsCommon.Open strSQL, adoCon
			
			'Place the forum ID's rs into array
			If NOT rsCommon.EOF Then iarryForumID = rsCommon.GetRows()
			
			'Close recordset
			rsCommon.Close
			
	
			'Becuase of issues with upgrades it's simpler to delete all permissions and have user re-create them
			strSQL = "DELETE FROM " & strDbTable & "Permissions;"
			
			'Excute delete query
			adoCon.Execute(strSQL)
			
			
			'Read in the groups
			strSQL = "SELECT " & strDbTable & "Group.Group_ID FROM " & strDbTable & "Group "
		
			'Query the database
			rsCommon.Open strSQL, adoCon
			
			'Place the date time data into an array
			If NOT rsCommon.EOF Then iaryGroupID = rsCommon.GetRows()
			
			'Relese server objects
			rsCommon.Close
			
			'If there are groups and forums populate the forum with group permisisons for each forum
			If isArray(iaryGroupID) AND isArray(iarryForumID) Then 
				
				'Loop through each forum
				Do While intForumLoopCounter <= Ubound(iarryForumID,2)
				
					'Reset permissions loop counter
					lngPermissionsLoopCounter = 0
				
					'Loop through topics and to populate the topics table
					Do While lngPermissionsLoopCounter <= Ubound(iaryGroupID,2)
					
						'Get the gorup ID
						blnRead = strDBTrue
						If CLng(iaryGroupID(0,lngPermissionsLoopCounter)) = 2 Then blnPost = strDBFalse Else blnPost = strDBTrue 'not guest
						If CLng(iaryGroupID(0,lngPermissionsLoopCounter)) = 2 Then blnReply = strDBFalse Else blnReply = strDBTrue 'not guest
						If CLng(iaryGroupID(0,lngPermissionsLoopCounter)) = 2 Then blnEdit = strDBFalse Else blnEdit = strDBTrue 'not guest
						If CLng(iaryGroupID(0,lngPermissionsLoopCounter)) = 2 Then blnDelete = strDBFalse Else blnDelete = strDBTrue 'not guest
						If CLng(iaryGroupID(0,lngPermissionsLoopCounter)) = 1 Then blnPriority = strDBTrue Else blnPriority = strDBFalse 'Admin only
						blnPollCreate = strDBFalse
						blnVote = strDBFalse
						blnAttachments = strDBFalse
						blnImageUpload = strDBFalse
						blnCheckFirst = strDBFalse
						blnEvents = strDBFalse
						blnModerator = strDBFalse
						
						'Insert values into database
						strSQL = "INSERT INTO [" & strDbTable & "Permissions] ( "
						strSQL = strSQL & "[Group_ID], "
						strSQL = strSQL & "[Forum_ID], "
						strSQL = strSQL & "[View_Forum], "
						strSQL = strSQL & "[Post], "
						strSQL = strSQL & "[Reply_posts], "
						strSQL = strSQL & "[Edit_posts], "
						strSQL = strSQL & "[Delete_posts], "
						strSQL = strSQL & "[Priority_posts], "
						strSQL = strSQL & "[Poll_create], "
						strSQL = strSQL & "[Vote], "
						strSQL = strSQL & "[Attachments], "
						strSQL = strSQL & "[Image_upload], "
						strSQL = strSQL & "[Moderate]"
						strSQL = strSQL & ") VALUES ("
						strSQL = strSQL & CLng(iaryGroupID(0,lngPermissionsLoopCounter)) & ", "			
						strSQL = strSQL & CLng(iarryForumID(0,intForumLoopCounter)) & ","
						strSQL = strSQL & blnRead & ","
						strSQL = strSQL & blnPost & ","
						strSQL = strSQL & blnReply & ","
						strSQL = strSQL & blnEdit & ","
						strSQL = strSQL & blnDelete & ","
						strSQL = strSQL & blnPriority & ","
						strSQL = strSQL & blnPollCreate & ","
						strSQL = strSQL & blnVote & ","
						strSQL = strSQL & blnAttachments & ","
						strSQL = strSQL & blnImageUpload & ","
						strSQL = strSQL & blnModerator & ")"
						
						'Write to the database
						adoCon.Execute(strSQL)
						
					
						'Move to next record
						lngPermissionsLoopCounter = lngPermissionsLoopCounter + 1
					Loop
					
					'Move to the next record
			                 intForumLoopCounter = intForumLoopCounter + 1
				Loop
			End If
			
			'Set all groups to have read access to all forums
			strSQL = "UPDATE " & strDbTable & "Permissions " & _
			"SET " & _
			strDbTable & "Permissions.View_Forum = " & strDBTrue & ";"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
	
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<strong>Complete</strong>';" & _
			vbCrLf & "</script>")
	
	
	
	
	
			'Stage 4 start
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br /><strong>Stage 4: Populating Database with Forum Statistics (this may take some time to complete)..... </strong>';" & _
			vbCrLf & "</script>")
	
			
	
			'Becuase Access doesn't populate new fields with the default value we need to do so now
			'Set the Last_post_author_ID to 1 or forums with no post will not display
			strSQL = "UPDATE " & strDbTable & "Forum " & _
			"SET " & _
			strDbTable & "Forum.Last_post_author_ID = 1;"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
	
			'Update Forum Table with last post details
			intForumLoopCounter = 0
			
			'Loop through each of the forums
			If isArray(iarryForumID) Then
				
				Do While intForumLoopCounter <= Ubound(iarryForumID,2)
				
					'Update forum stats using function
					updateForumStats(CInt(iarryForumID(0,intForumLoopCounter)))
				
					 'Move to the next record
			                 intForumLoopCounter = intForumLoopCounter + 1
				Loop
			End If
			
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<strong>Complete</strong>';" & _
			vbCrLf & "</script>")
	
	
	
	
	
			'Update Topic Table with first, last, post details and no. of replies
			
			'Stage 5 start
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br /><strong>Stage 5: Populating Database with Topic Statistics (this may take some time to complete)..... </strong>';" & _
			vbCrLf & "</script>")
			
			
			
			
			'Get the topic ID's from the database
			strSQL = "SELECT " & strDbTable & "Topic.Topic_ID " & _
			"FROM " & strDbTable & "Topic " & _
			"ORDER BY " & strDbTable & "Topic.Topic_ID ASC;"
			
			'Get rs
			rsCommon.Open strSQL, adoCon
			
			'Place the recordset into an array for improved performance
			If NOT rsCommon.EOF Then larryTopicID = rsCommon.GetRows()
			
			'Close rs
			rsCommon.Close
			
			If isArray(larryTopicID) Then
			
				'Loop through topics and to populate the topics table
				Do While lngTopicLoopCounter <= Ubound(larryTopicID,2)
				
					'Read in the topic thread to update the topics stats in the topic table
					strSQL = "SELECT " & strDbTable & "Thread.Thread_ID " & _
					"FROM " & strDbTable & "Thread " & _
					"WHERE " & strDbTable & "Thread.Topic_ID = " & CLng(larryTopicID(0,lngTopicLoopCounter)) & ";"
					
					
					'Open recordset
					rsCommon.Open strSQL, adoCon, 1,2
					
					'If a record returned update the database
					If NOT rsCommon.EOF Then
						
						intNoOfReplies = Cint(rsCommon.RecordCount)-1
						lngStartThreadID = CLng(rsCommon("Thread_ID"))
						
						'Move last record
						rsCommon.MoveLast
						
						lngLastThreadID = CLng(rsCommon("Thread_ID"))
						
						'Update the database
						strSQL = "UPDATE " & strDbTable & "Topic " & _
						"SET " & strDbTable & "Topic.Start_Thread_ID = " & lngStartThreadID & ", " & _
						strDbTable & "Topic.Last_Thread_ID = " & lngLastThreadID & ", " & _
						strDbTable & "Topic.No_of_replies = " & intNoOfReplies & " " & _
						"WHERE " & strDbTable & "Topic.Topic_ID = " & CLng(larryTopicID(0,lngTopicLoopCounter)) & ";"
						
						'Write to the database
						adoCon.Execute(strSQL)
					End If
					
					'Close recordset
					rsCommon.Close
					
					'Move next array position
					lngTopicLoopCounter = lngTopicLoopCounter + 1
				Loop
	
			End If
			
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<strong>Complete</strong>';" & _
			vbCrLf & "</script>")
			
	
	
	
			
			'Display a message to say the database is created
			If blnErrorOccured = True Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />" & Err.description & "<br /><br /><h2>Access database is updated to 9.x, but with Error!</h2>'" & _
				vbCrLf & "</script>")
			End If
			
			
			
		End If
	End If

	'Reset Server Variables
	Set adoCon = Nothing
End If

%>