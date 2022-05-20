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
If Request.QueryString("setup") = "SqlServer9Update" Then
	Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
	vbCrLf & "	document.getElementById('displayState').innerHTML = 'Your Database is being updated. Please be patient as this may take a few minutes to complete.';" & _
	vbCrLf & "</script>")

Else
	Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
	vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br /><br /><strong>Part 2: Version 9.x to 10.x Database Update.</strong> Please be patient as this may take a few minutes to complete.';" & _
	vbCrLf & "</script>")

End If




'Resume on all errors
On Error Resume Next

'intialise variables
blnErrorOccured = False


'If a username and password is entred then start the ball rolling
If strDatabaseType = "SQLServer" AND strSQLDBUserName <> "" Then
	
	
	'Open the database
	Call openDatabase(strCon)

	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then
		
		
		
		Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
		vbCrLf & "	document.getElementById('displayState').innerHTML = '<img src=""forum_images/error.png"" alt=""Error"" /> <strong>Error</strong><br /><strong>Error Connecting to database on SQL Server</strong><br /><br />Replace the database/database_settings.asp file with the one from the orginal Web Wiz Forums download and start the setup process again.<br /><br /><strong>Error Details:</strong><br />" & Err.description & "';" & _
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
		
		
		
		
		'Check to see if the database has been updated from 8.x
		
		'the info field is setup in for 9.x
		strSQL = "SELECT " & strDbTable & "Author.Gender " & _
		"FROM " & strDbTable & "Author " & _
		"WHERE " & strDbTable & "Author.Author_ID = 1;"
		
		'Query the database
		rsCommon.Open strSQL, adoCon
		
		'If error occured the database has not been updated to 7.x
		If NOT CLng(Err.Number) = 0 Then
			
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "	document.getElementById('displayState').innerHTML = '<img src=""forum_images/error.png"" alt=""Error"" /> <strong>Error</strong><br /><strong>The Database Setup Wizard has detected that the database you are updating is not a Web Wiz Forums version 9.x database.</strong><br /><br />Replace the database/database_settings.asp file with the one from the orginal Web Wiz Forums download and start the setup process again.';" & _
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
		strSQL = "SELECT " & strDbTable & "Author.Points " & _
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
			
			
			'On Error goto 0


'******************************************
'***  	Update/Create the tables      *****
'******************************************

			'Stage one start
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br /><strong>Stage 1: Creating New Database Tables and Fields..... </strong>';" & _
			vbCrLf & "</script>")
			
			
			
			
			'Create the SetupOptions Table
			strSQL = "CREATE TABLE [" & strDBO & "].[" & strDbTable & "SetupOptions] ("
			strSQL = strSQL & "[Option_Item] [nvarchar] (30),"
			strSQL = strSQL & "[Option_Value] [text] NULL " 
			strSQL = strSQL & ")"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br />Error Creating the Table " & strDbTable & "SetupOptions <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			'Create the LadderGroup Table
			strSQL = "CREATE TABLE [" & strDBO & "].[" & strDbTable & "LadderGroup] ("
			strSQL = strSQL & "[Ladder_ID] [int] IDENTITY (1, 1) PRIMARY KEY  CLUSTERED  NOT NULL ,"
			strSQL = strSQL & "[Ladder_Name] [nvarchar] (25) " 
			strSQL = strSQL & ")"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br />Error Creating the Table " & strDbTable & "LadderGroup <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			'Create the Spam Table
			strSQL = "CREATE TABLE [" & strDBO & "].[" & strDbTable & "Spam] ("
			strSQL = strSQL & "[Spam_ID] [int] IDENTITY (1, 1) PRIMARY KEY  CLUSTERED  NOT NULL ,"
			strSQL = strSQL & "[Spam] [nvarchar] (255), " 
			strSQL = strSQL & "[Spam_Action] [nvarchar] (20) " 
			strSQL = strSQL & ")"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br />Error Creating the Table " & strDbTable & "Spam <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			'Create the TopicRatingVote Table
			strSQL = "CREATE TABLE [" & strDBO & "].[" & strDbTable & "TopicRatingVote] ("
			strSQL = strSQL & "[Topic_ID] [int] NOT NULL DEFAULT (0), "
			strSQL = strSQL & "[Author_ID] [int] NOT NULL DEFAULT (0) " 
			strSQL = strSQL & ")"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br />Error Creating the Table " & strDbTable & "TopicRatingVote <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			'Create the ThreadThanks Table
			strSQL = "CREATE TABLE [" & strDBO & "].[" & strDbTable & "ThreadThanks] ("
			strSQL = strSQL & "[Thread_ID] [int] NOT NULL DEFAULT (0),"
			strSQL = strSQL & "[Author_ID] [int] NOT NULL DEFAULT (0) " 
			strSQL = strSQL & ")"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br />Error Creating the Table " & strDbTable & "ThreadThanks <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			
			'Update tblAuthor
			strSQL = "ALTER TABLE [" & strDBO & "].[" & strDbTable & "Author] ADD "
			strSQL = strSQL & "[Points] [int] NOT NULL DEFAULT (0) ,"
			strSQL = strSQL & "[Login_IP] [nvarchar] (50) NULL ,"
			strSQL = strSQL & "[Custom1] [nvarchar] (255) NULL ,"
			strSQL = strSQL & "[Custom2] [nvarchar] (255) NULL ,"
			strSQL = strSQL & "[Custom3] [nvarchar] (255) NULL ,"
			strSQL = strSQL & "[Answered] [int] NOT NULL DEFAULT (0) ,"
			strSQL = strSQL & "[Thanked] [int] NOT NULL DEFAULT (0) ,"
			strSQL = strSQL & "[LinkedIn] [nvarchar] (75) NULL ,"
			strSQL = strSQL & "[Facebook] [nvarchar] (75) NULL ,"
			strSQL = strSQL & "[Twitter] [nvarchar] (75) NULL ,"
			strSQL = strSQL & "[Inbox_no_of_PM] [int] DEFAULT (0) "
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br />Error Altering the Table " & strDbTable & "Group <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			'Update tblForum
			strSQL = "ALTER TABLE [" & strDBO & "].[" & strDbTable & "Forum] ADD "
			strSQL = strSQL & "[Forum_URL] [nvarchar] (80) NULL ,"
			strSQL = strSQL & "[Forum_icon] [nvarchar] (80) NULL "
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br />Error Altering the Table " & strDbTable & "Forum <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			'Update tblTopic
			strSQL = "ALTER TABLE [" & strDBO & "].[" & strDbTable & "Topic] ADD "
			strSQL = strSQL & "[Rating] [float] NOT NULL DEFAULT (0) ,"
			strSQL = strSQL & "[Rating_Total] [int] NOT NULL DEFAULT (0) ,"
			strSQL = strSQL & "[Rating_Votes] [int] NOT NULL DEFAULT (0) "
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br />Error Altering the Table " & strDbTable & "Topic <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			'Update tblGroup
			strSQL = "ALTER TABLE [" & strDBO & "].[" & strDbTable & "Group] ADD "
			strSQL = strSQL & "[Ladder_ID] [int] NOT NULL DEFAULT (0) ,"
			strSQL = strSQL & "[Signatures] [bit] NOT NULL DEFAULT (1), "
			strSQL = strSQL & "[URLs] [bit] NOT NULL DEFAULT (1), "
			strSQL = strSQL & "[Images] [bit] NOT NULL DEFAULT (1), "
			strSQL = strSQL & "[Private_Messenger] [bit] NOT NULL DEFAULT (1), "
			strSQL = strSQL & "[Chat_Room] [bit] NOT NULL DEFAULT (1) "
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br />Error Altering the Table " & strDbTable & "Group <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			'Update PMMessage
			strSQL = "ALTER TABLE [" & strDBO & "].[" & strDbTable & "PMMessage] ADD "
			strSQL = strSQL & "[Inbox] [bit] NOT NULL DEFAULT (1), "
			strSQL = strSQL & "[Outbox] [bit] NOT NULL DEFAULT (1) "
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br />Error Altering the Table " & strDbTable & "PMMessage <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			'Update BanList
			strSQL = "ALTER TABLE [" & strDBO & "].[" & strDbTable & "BanList] ADD "
			strSQL = strSQL & "[Reason] [nvarchar] (50) NULL "
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br />Error Altering the Table " & strDbTable & "BanList <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			'Update Thread
			'Done in two parts for larger databases
			strSQL = "ALTER TABLE [" & strDBO & "].[" & strDbTable & "Thread] ADD "
			strSQL = strSQL & "[Answer] [bit] NOT NULL DEFAULT (0) "
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br />Error Altering the Table " & strDbTable & "Thread.Answer <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
	
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<strong>Complete</strong>';" & _
			vbCrLf & "</script>")
			
	
			
			'Update Thread
			strSQL = "ALTER TABLE [" & strDBO & "].[" & strDbTable & "Thread] ADD "
			strSQL = strSQL & "[Thanks] [int] NOT NULL DEFAULT (0) "
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br />Error Altering the Table " & strDbTable & "Thread.Thanks <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			'Altering Thread
			strSQL = "ALTER TABLE tblThread ALTER COLUMN [IP_addr] nvarchar(50)"
			
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
	
	
			'Intialise the main ADO recordset object
			Set rsCommon = Server.CreateObject("ADODB.Recordset")
			Set rsCommon2 = Server.CreateObject("ADODB.Recordset")
			
			
			'Enter the default values in the LadderGroup Table
			'Primary Ladder Group
			strSQL = "INSERT INTO [" & strDBO & "].[" & strDbTable & "LadderGroup] ("
			strSQL = strSQL & "[Ladder_Name] "
			strSQL = strSQL & ") "
			strSQL = strSQL & "VALUES "
			strSQL = strSQL & "('Primary Ladder Group')"
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br />Error entering default values in the Table " & strDbTable & "LadderGroup<br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
	
			
			'Enter the default values in the Group Table
			strSQL = "UPDATE " & strDbTable & "Group " & _
			"SET " & _
			strDbTable & "Group.Ladder_ID = 1, " & _
			strDbTable & "Group.Signatures = " & strDBTrue & ", " & _
			strDbTable & "Group.URLs = " & strDBTrue & ", " & _
			strDbTable & "Group.Images = " & strDBTrue & ", " & _
			strDbTable & "Group.Private_Messenger = " & strDBTrue & ", " & _
			strDbTable & "Group.Chat_Room = " & strDBTrue & "; "
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br />Error entering default values in the Table " & strDbTable & "Group<br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			'Enter default values for PM inbox and outbox
			strSQL = "UPDATE " & strDbTable & "PMMessage " & _
			"SET " & _
			strDbTable & "PMMessage.Inbox = " & strDBTrue & ", " & _
			strDbTable & "PMMessage.Outbox = " & strDBTrue & "; "
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br />Error entering default values in the Table " & strDbTable & "PMMessage<br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
	
	
			
			
			'Set default configuration values
			Call addConfigurationItem("A_code", "True")
			Call addConfigurationItem("Active_users", "True")
			Call addConfigurationItem("Active_users_viewing", "True")
			Call addConfigurationItem("Answer_posts", "admin")
			Call addConfigurationItem("Answer_wording", "Answer")
			Call addConfigurationItem("Avatar", "True")
			Call addConfigurationItem("Calendar", "True")
			Call addConfigurationItem("CAPTCHA", "True")
			Call addConfigurationItem("Chat_room", "True")
			Call addConfigurationItem("Cookie_domain", "")
			Call addConfigurationItem("Cookie_path", "/")
			Call addConfigurationItem("Cookie_prefix", "wwf10")
			Call addConfigurationItem("Cust_item_name_1", "")
			Call addConfigurationItem("Cust_item_name_2", "")
			Call addConfigurationItem("Cust_item_name_3", "")
			Call addConfigurationItem("Cust_item_name_req_1", "True")
			Call addConfigurationItem("Cust_item_name_req_2", "True")
			Call addConfigurationItem("Cust_item_name_req_3", "True")
			Call addConfigurationItem("Cust_item_name_view_1", "True")
			Call addConfigurationItem("Cust_item_name_view_2", "True")
			Call addConfigurationItem("Cust_item_name_view_3", "True")
			Call addConfigurationItem("Date_today_bold", "True")
			Call addConfigurationItem("Detailed_error_reporting", "False")
			Call addConfigurationItem("Edit_post_time_frame", "0")
			Call addConfigurationItem("Edited_by_delay", "0")
			Call addConfigurationItem("Email_activate", "False")
			Call addConfigurationItem("Email_all_notifications", "False")
			Call addConfigurationItem("email_notify", "True")
			Call addConfigurationItem("Email_post", "True")
			Call addConfigurationItem("Email_sys", "True")
			Call addConfigurationItem("Emoticons", "True")
			Call addConfigurationItem("Facebook_image", "")
			Call addConfigurationItem("Facebook_likes", "True")
			Call addConfigurationItem("Facebook_page_ID", "")
			Call addConfigurationItem("Flash", "True")
			Call addConfigurationItem("Footer", "</body>" & vbCrLf & "</html>")
			Call addConfigurationItem("Footer_mobile", "</body>" & vbCrLf & "</html>")
			Call addConfigurationItem("Form_CAPTCHA", "True")
			Call addConfigurationItem("forum_email_address", "forum@example.com")
			Call addConfigurationItem("Forum_header_ad", "")
			Call addConfigurationItem("Forums_message", "")
			Call addConfigurationItem("forum_name", "Web Wiz Forums")
			Call addConfigurationItem("forum_path", "http://www.example.com")
			Call addConfigurationItem("Forum_post_ad", "")
			Call addConfigurationItem("Forums_closed", "False")
			Call addConfigurationItem("Google_plus_1", "True")
			Call addConfigurationItem("Guest_SID", "False")
			Call addConfigurationItem("Header", "</head>" & vbCrLf & "<body>")
			Call addConfigurationItem("Header_mobile", "</head>" & vbCrLf & "<body>")
			Call addConfigurationItem("Homepage", "True")
			Call addConfigurationItem("Hot_replies", "20")
			Call addConfigurationItem("Hot_views", "100")
			Call addConfigurationItem("Hyperlinks_nofollow", "True")
			Call addConfigurationItem("IE_editor", "True")
			Call addConfigurationItem("Install_ID", "Adware")
			Call addConfigurationItem("L_code", "True")
			Call addConfigurationItem("Location", "False")
			Call addConfigurationItem("Login_attempts", "3")
			Call addConfigurationItem("Long_reg", "True")
			Call addConfigurationItem("mail_component", "CDOSYS")
			Call addConfigurationItem("Mail_password", "")
			Call addConfigurationItem("mail_server", "localhost")
			Call addConfigurationItem("Mail_server_port", "25")
			Call addConfigurationItem("Mail_username", "")
			Call addConfigurationItem("Member_approve", "False")
			Call addConfigurationItem("Member_Profile_View", "members")
			Call addConfigurationItem("Meta_description", "This is a forum powered by Web Wiz Forums. To find out about Web Wiz Forums, go to www.WebWizForums.com")
			Call addConfigurationItem("Meta_keywords", "community,forums,chat,talk,discussions")
			Call addConfigurationItem("Meta_tags_dynamic", "True")
			Call addConfigurationItem("Min_password_length", "5")
			Call addConfigurationItem("Min_usename_length", "3")
			Call addConfigurationItem("Mobile_View", "True")
			Call addConfigurationItem("Mod_profile_edit", "True")
			Call addConfigurationItem("Most_active_date", "2011-01-01 00:00:00")
			Call addConfigurationItem("Most_active_users", "5")
			Call addConfigurationItem("NewsPad", "True")
			Call addConfigurationItem("NewsPad_URL", "")
			Call addConfigurationItem("Page_encoding", "utf-8")
			Call addConfigurationItem("Password_complexity", "False")
			Call addConfigurationItem("PM_Flash", "True")
			Call addConfigurationItem("PM_Flood", "5")
			Call addConfigurationItem("PM_inbox", "100")
			Call addConfigurationItem("PM_outbox", "150")
			Call addConfigurationItem("PM_overusage_action", "delete")
			Call addConfigurationItem("PM_spam_ignore", "True")
			Call addConfigurationItem("PM_YouTube", "True")
			Call addConfigurationItem("Points_answer", "5")
			Call addConfigurationItem("Points_reply", "1")
			Call addConfigurationItem("Points_thanked", "5")
			Call addConfigurationItem("Points_topic", "2")
			Call addConfigurationItem("Post_order", "ASC")
			Call addConfigurationItem("Post_thanks", "True")
			Call addConfigurationItem("Private_msg", "True")
			Call addConfigurationItem("Process_time", "True")
			Call addConfigurationItem("Quick_reply", "True")
			Call addConfigurationItem("Real_name", "False")
			Call addConfigurationItem("Reg_closed", "False")
			Call addConfigurationItem("Registration_Rules", "<p>If you agree with the following rules then click on the 'Accept' button at the bottom of the page if not click on the 'Cancel' button.</p>" & vbCrLf & vbCrLf & "<p>When you register you are required to give a small amount of information, much of which is optional, anything you do give must be considered as becoming public information.</p>" & vbCrLf & vbCrLf & "<p>You agree not to use this forum to post any material which is vulgar, defamatory, inaccurate, harassing, hateful, threatening, invading of others privacy, sexually oriented, or violates any laws. You also agree that you will not post any copyrighted material that is not owned by yourself or the owners of these forums.</p><p>" & vbCrLf & vbCrLf & "You remain solely responsible for the content of your messages, and you agree to indemnify and hold harmless this forum and their agents with respect to any claim based upon any post you may make. We also reserve the right to reveal whatever information we know about you in the event of a complaint or legal action arising from any message posted by yourself.</p>" & vbCrLf & vbCrLf & "<p>Although messages posted are not the responsibility of this forum and we are not responsible for the content or accuracy of any of these messages, we reserve the right to delete any message for any or no reason whatsoever. If you do find any posts are objectionable then please contact the forum by e-mail.</p>" & vbCrLf & vbCrLf & "<p>By posting content on this forum you are accepting that you grant full rights to the publication of your message content within the forum and that you can not later withdraw publication rights, including but not limited to the removal, editing, and/or modifying of the published content.</p>" & vbCrLf & vbCrLf & "<p>The Federal Trade Commission's Children's Online Privacy Protection Act of 1998 (COPPA) requires that Web Sites are to obtain parental consent before collecting, using, or disclosing personal information from children under 13. <b>If you are below 13 then you can NOT use this forum. Do NOT register if you are below the age of 13.</b></p>" & vbCrLf & vbCrLf & "<p>This forum uses cookies to identify logged in members as they move around or return to the forum. Cookies are also used to enhance the forum service, for example to remember forum options and display new posts since your last visit.</p>" & vbCrLf & vbCrLf & "<p>By registering to use this forum you meet the above criteria and agree to abide by all of the above rules and policies.<br /></p>")
			Call addConfigurationItem("RSS", "True")
			Call addConfigurationItem("RSS_max_results", "10")
			Call addConfigurationItem("RSS_TTL", "30")
			Call addConfigurationItem("Search_eng_sessions", "False")
			Call addConfigurationItem("Search_time_default", "6")
			Call addConfigurationItem("SEO_title", "True")
			Call addConfigurationItem("Session_db", "True")
			Call addConfigurationItem("Share_topics_links", "True")
			Call addConfigurationItem("Show_birthdays", "True")
			Call addConfigurationItem("Show_edit", "True")
			Call addConfigurationItem("Show_Forum_Stats", "True")
			Call addConfigurationItem("Show_latest_posts", "True")
			Call addConfigurationItem("Show_Member_list", "True")
			Call addConfigurationItem("Show_mod", "True")
			Call addConfigurationItem("Show_todays_birthdays", "True")
			Call addConfigurationItem("Show_header_footer", "False")
			Call addConfigurationItem("Show_mobile_header_footer", "False")
			Call addConfigurationItem("Signatures", "True")
			Call addConfigurationItem("Skin_file", "css_styles/default/")
			Call addConfigurationItem("Skin_image_path", "forum_images/")
			Call addConfigurationItem("Skin_nav_spacer", " > ")
			Call addConfigurationItem("Spam_minutes", "10")
			Call addConfigurationItem("Spam_seconds", "30")
			Call addConfigurationItem("Text_direction", "ltr")
			Call addConfigurationItem("Text_link", "False")
			Call addConfigurationItem("Threads_per_page", "10")
			Call addConfigurationItem("Title_image", "forum_images/web_wiz_forums.png")
			Call addConfigurationItem("Topic_icon", "True")
			Call addConfigurationItem("Topic_rating", "True")
			Call addConfigurationItem("Topics_new_bold", "True")
			Call addConfigurationItem("Topics_per_page", "24")
			Call addConfigurationItem("Tracking_code_update", "False")
			Call addConfigurationItem("Twitter_tweet", "True")
			Call addConfigurationItem("URL_Rewriting", "False")
			Call addConfigurationItem("Upload_allocation", "10")
			Call addConfigurationItem("Upload_avatar", "True")
			Call addConfigurationItem("Upload_avatar_size", "100")
			Call addConfigurationItem("Upload_avatar_types", "jpg;jpeg;gif;png")
			Call addConfigurationItem("Upload_component", "AspUpload")
			Call addConfigurationItem("Upload_file_scan", "False")
			Call addConfigurationItem("Upload_files_size", "1024")
			Call addConfigurationItem("Upload_files_type", "zip;rar;doc;pdf;txt;rtf;gif;jpg;png;mp3;docx;xls;xlsx")
			Call addConfigurationItem("Upload_img_size", "100")
			Call addConfigurationItem("Upload_img_types", "jpg;jpeg;gif;png")
			Call addConfigurationItem("VigLink_key", "")
			Call addConfigurationItem("Vote_choices", "7")
			Call addConfigurationItem("website_name", "My Website")
			Call addConfigurationItem("website_path", "http://www.webwizforums.com")
			Call addConfigurationItem("YouTube", "True")
			Call addConfigurationItem("HTTP_XML_API", "False")
			Call addConfigurationItem("SFSpam", "False")
			Call addConfigurationItem("SFSpam_Agreement", "")
			Call addConfigurationItem("SFSpam_API_Key", "")
			Call addConfigurationItem("SFSpam_no_blocked", "")
			Call addConfigurationItem("SFSpam_last_block_date", "")
			Call addConfigurationItem("SFSpam_last_block_email", "")
			Call addConfigurationItem("SFSpam_last_block_IP", "")
			Call addConfigurationItem("SFSpam_Username", "False")
			Call addConfigurationItem("SSL", "Disabled")
			Call addConfigurationItem("EU_Cookie_Law", "False")
			Call addConfigurationItem("Local_links_window", "True")
			Call addConfigurationItem("Registration_CAPTCHA", "True")
			Call addConfigurationItem("Session_Timeout", "20")
			Call addConfigurationItem("Session_IP_Check", "True")
			Call addConfigurationItem("Upload_Duplicate_Files", "True")
			
			
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br />Error creating default configuration<br />" & Err.description & ".';" & _
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
	'***  	Migrate default values	      *****
	'******************************************
			
			
			'Stage 3 start
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br /><strong>Stage 3: Migrating forum configuration settings to new format..... </strong>';" & _
			vbCrLf & "</script>")
		
			
			'Read in old values from configuration table and update the database with them
			'Initialise the SQL variable with an SQL statement to get the configuration details from the database
			strSQL = "SELECT " & strDbTable & "Configuration.* " & _
			"FROM " & strDbTable & "Configuration" &  strDBNoLock & " " & _
			"WHERE " & strDbTable & "Configuration.ID = 1;"
		
			'Query the database
			rsCommon2.Open strSQL, adoCon
			
			
			'If there is config deatils in the recordset then read them in
			If NOT rsCommon2.EOF Then
				
				Call addConfigurationItem("website_name", rsCommon2("website_name"))
				Call addConfigurationItem("forum_name", rsCommon2("forum_name"))
				Call addConfigurationItem("website_path", rsCommon2("website_path"))
				Call addConfigurationItem("forum_path", rsCommon2("forum_path"))
				Call addConfigurationItem("mail_component", rsCommon2("mail_component"))
				Call addConfigurationItem("mail_server", rsCommon2("mail_server"))
				Call addConfigurationItem("forum_email_address", rsCommon2("forum_email_address"))
				Call addConfigurationItem("email_notify", CBool(rsCommon2("email_notify")))
				Call addConfigurationItem("Text_link", rsCommon2("Text_link"))
				Call addConfigurationItem("IE_editor", CBool(rsCommon2("IE_editor")))
				Call addConfigurationItem("Topics_per_page", CInt(rsCommon2("Topics_per_page")))
				Call addConfigurationItem("Title_image", rsCommon2("Title_image"))
				Call addConfigurationItem("Emoticons", CBool(rsCommon2("Emoticons")))
				Call addConfigurationItem("Avatar", CBool(rsCommon2("Avatar")))
				Call addConfigurationItem("Email_activate", CBool(rsCommon2("Email_activate")))
			 	Call addConfigurationItem("Hot_views", CInt(rsCommon2("Hot_views")))
				Call addConfigurationItem("Hot_replies", CInt(rsCommon2("Hot_replies")))
				Call addConfigurationItem("Email_post", CBool(rsCommon2("Email_post")))
				Call addConfigurationItem("Private_msg", CBool(rsCommon2("Private_msg")))
				Call addConfigurationItem("No_of_priavte_msg", CInt(rsCommon2("No_of_priavte_msg")))
				Call addConfigurationItem("Threads_per_page", CInt(rsCommon2("Threads_per_page")))
				Call addConfigurationItem("Spam_seconds", CInt(rsCommon2("Spam_seconds")))
				Call addConfigurationItem("Spam_minutes", CInt(rsCommon2("Spam_minutes")))
				Call addConfigurationItem("Vote_choices", CInt(rsCommon2("Vote_choices")))
				Call addConfigurationItem("Email_sys", CBool(rsCommon2("Email_sys")))
				Call addConfigurationItem("Active_users", CBool(rsCommon2("Active_users")))
				Call addConfigurationItem("Forums_closed", CBool(rsCommon2("Forums_closed")))
				Call addConfigurationItem("Show_edit", CBool(rsCommon2("Show_edit")))
				Call addConfigurationItem("Process_time", CBool(rsCommon2("Process_time")))
				Call addConfigurationItem("Flash", CBool(rsCommon2("Flash")))
				Call addConfigurationItem("Show_mod", CBool(rsCommon2("Show_mod")))
				Call addConfigurationItem("Upload_avatar", CBool(rsCommon2("Upload_avatar")))
				Call addConfigurationItem("Reg_closed", CBool(rsCommon2("Reg_closed")))
				Call addConfigurationItem("Upload_component", rsCommon2("Upload_component"))
				Call addConfigurationItem("Upload_img_types", rsCommon2("Upload_img_types"))
				Call addConfigurationItem("Upload_img_size", CInt(rsCommon2("Upload_img_size")))
				Call addConfigurationItem("Upload_files_type", rsCommon2("Upload_files_type"))
				Call addConfigurationItem("Upload_files_size", CInt(rsCommon2("Upload_files_size")))
				Call addConfigurationItem("Upload_allocation", CInt(rsCommon2("Upload_allocation")))
				Call addConfigurationItem("Mail_username", rsCommon2("Mail_username"))
				Call addConfigurationItem("Mail_password", rsCommon2("Mail_password"))
				Call addConfigurationItem("Skin_file", rsCommon2("Skin_file"))
				Call addConfigurationItem("Skin_image_path", rsCommon2("Skin_image_path"))
				Call addConfigurationItem("Skin_nav_spacer", rsCommon2("Skin_nav_spacer"))
				Call addConfigurationItem("Topic_icon", CBool(rsCommon2("Topic_icon")))
				Call addConfigurationItem("Long_reg", CBool(rsCommon2("Long_reg")))
				Call addConfigurationItem("CAPTCHA", CBool(rsCommon2("CAPTCHA")))
				Call addConfigurationItem("Calendar", CBool(rsCommon2("Calendar")))
				Call addConfigurationItem("Guest_SID", CBool(rsCommon2("Guest_SID")))
				Call addConfigurationItem("Member_approve", CBool(rsCommon2("Member_approve")))
				Call addConfigurationItem("RSS", CBool(rsCommon2("RSS")))	
				Call addConfigurationItem("PM_Flood", CInt(rsCommon2("PM_Flood")))
				Call addConfigurationItem("NewsPad", CBool(rsCommon2("NewsPad")))
				Call addConfigurationItem("NewsPad_URL", rsCommon2("NewsPad_URL"))
			End If
			
			'Close recordset
			rsCommon2.Close
			
			
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br />Error migrating configuration settings to new format<br />" & Err.description & ".';" & _
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
	'***  	Migrate posts to points	      *****
	'******************************************
			
			
			'Stage 3 start
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br /><strong>Stage 4: Migrate Post Count to new Point System (this may take some time to complete)..... </strong>';" & _
			vbCrLf & "</script>")	
			
			
			
			'Initliase the SQL query to get all the posts in this topic that are not hidden
			strSQL = "SELECT" & " " & strDbTable & "Author.Author_ID,  " & strDbTable & "Author.No_of_posts " & _
			"FROM " & strDbTable & "Author " & _
			"WHERE " & strDbTable & "Author.No_of_posts > 0;"
			
			'Query the database
			rsCommon2.Open strSQL, adoCon
			
			'Place the recordset into an array for improved performance
			If NOT rsCommon2.EOF Then sarryNoPosts = rsCommon2.GetRows()
						
			'Close rs
			rsCommon2.Close
			
			'If we have an array loop through it			
			If isArray(sarryNoPosts) Then
			
				'Loop through update points
				Do While lngPostsLoopCounter <= Ubound(sarryNoPosts,2)
				
					'See if posts is above 0
					If CLng(sarryNoPosts(1,lngPostsLoopCounter)) > 0 Then
				
						strSQL = "UPDATE " & strDbTable & "Author " & strRowLock & " " & _
						"SET " & strDbTable & "Author.Points = " & CLng(sarryNoPosts(1,lngPostsLoopCounter)) & " " & _
						"WHERE " & strDbTable & "Author.Author_ID = " & CLng(sarryNoPosts(0,lngPostsLoopCounter)) & ";"
						
									
						'Write the updated number of posts to the database
						adoCon.Execute(strSQL)
					
					End If
				
					'Move next record
					lngPostsLoopCounter = lngPostsLoopCounter + 1
				Loop
			End If
			
			
			
			'Close recordset
			Set rsCommon = Nothing
			Set rsCommon2 = Nothing
			
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
End If
%>
      