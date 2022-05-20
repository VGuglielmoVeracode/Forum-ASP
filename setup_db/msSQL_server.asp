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
Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
vbCrLf & "	document.getElementById('displayState').innerHTML = 'Your Database is being created. Please be patient as this may take a few minutes to complete.';" & _
vbCrLf & "</script>")
	




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
		
		
		'Check to see if the database is already created
		
		'Intialise the main ADO recordset object
		Set rsCommon = CreateObject("ADODB.Recordset")
		
		'Get the admin account
		strSQL = "SELECT " & strDbTable & "Author.Username " & _
		"FROM " & strDbTable & "Author " & _
		"WHERE " & strDbTable & "Author.Author_ID = 1;"
		
		'Query the database
		rsCommon.Open strSQL, adoCon
		
		
		'If no error occured the database has been created
		If CLng(Err.Number) = 0 Then
			
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "	document.getElementById('displayState').innerHTML = '<img src=""forum_images/error.png"" alt=""Error"" /> <strong>Error</strong><br /><strong>The Database Setup Wizard has detected that your database has already been created.</strong><br /><br />Click here to go to your <a href=""default.asp"">Web Wiz Forum Homepage</a>.';" & _
			vbCrLf & "</script>")
			
			
			Set rsCommon = Nothing
		
		'Create the database
		Else

			'Reset error object
			Err.Number = 0
			Set rsCommon = Nothing
			
			
			'For testing
			'On Error goto 0



'******************************************
'***  		Create the tables     *****
'******************************************


			'Stage one start
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br /><strong>Stage 1: Creating New Database Tables and Fields..... </strong>';" & _
			vbCrLf & "</script>")


			'Create the Category Table
			strSQL = "CREATE TABLE [" & strDBO & "].[" & strDbTable & "Category] ("
			strSQL = strSQL & "[Cat_ID] [smallint] IDENTITY (1, 1) PRIMARY KEY  CLUSTERED  NOT NULL ,"
			strSQL = strSQL & "[Cat_name] [nvarchar] (60) NOT NULL ,"
			strSQL = strSQL & "[Cat_order] [smallint] NOT NULL DEFAULT (1) "
			strSQL = strSQL & ") ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
	
				'Write an error message
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "Category <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
	
	
			'Create the Forum Table
			strSQL = "CREATE TABLE [" & strDBO & "].[" & strDbTable & "Forum] ("
			strSQL = strSQL & "[Forum_ID] [int] IDENTITY (1, 1) PRIMARY KEY  CLUSTERED  NOT NULL ,"
			strSQL = strSQL & "[Cat_ID] [smallint] NULL ,"
			strSQL = strSQL & "[Sub_ID] [smallint] NOT NULL DEFAULT (0),"
			strSQL = strSQL & "[Forum_Order] [smallint] NOT NULL DEFAULT (0),"
			strSQL = strSQL & "[Forum_name] [nvarchar] (70) NULL ,"
			strSQL = strSQL & "[Forum_description] [nvarchar] (200) NULL ,"
			strSQL = strSQL & "[No_of_topics] [int] NOT NULL DEFAULT (0),"
			strSQL = strSQL & "[No_of_posts] [int] NOT NULL DEFAULT (0),"
			strSQL = strSQL & "[Last_post_author_ID] [int] NOT NULL DEFAULT (1),"
			strSQL = strSQL & "[Last_post_date] [datetime] NULL,"
			strSQL = strSQL & "[Last_topic_ID] [int]  DEFAULT (0), "
			strSQL = strSQL & "[Password] [nvarchar] (50) NULL ,"
			strSQL = strSQL & "[Forum_code] [nvarchar] (40) NULL ,"
			strSQL = strSQL & "[Show_topics] [smallint] NULL, "
			strSQL = strSQL & "[Locked] [bit] NOT NULL ,"
			strSQL = strSQL & "[Hide] [bit] NOT NULL, "
			strSQL = strSQL & "[Forum_URL] [nvarchar] (80) NULL ,"
			strSQL = strSQL & "[Forum_icon] [nvarchar] (80) NULL "
			strSQL = strSQL & ") ON [PRIMARY]"
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
	
				'Write an error message
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "Forum <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
	
	
			'Create the Topic Table
			strSQL = "CREATE TABLE [" & strDBO & "].[" & strDbTable & "Topic] ("
			strSQL = strSQL & "[Topic_ID] [int] IDENTITY (1, 1) PRIMARY KEY  CLUSTERED  NOT NULL ,"
			strSQL = strSQL & "[Forum_ID] [int] NULL ,"
			strSQL = strSQL & "[Poll_ID] [int] NOT NULL DEFAULT (0),"
			strSQL = strSQL & "[Moved_ID] [int] NOT NULL DEFAULT (0),"
			strSQL = strSQL & "[Subject] [nvarchar] (70) NOT NULL,"
			strSQL = strSQL & "[Icon] [nvarchar] (20) NULL,"
			strSQL = strSQL & "[Start_Thread_ID] [int] NULL,"
			strSQL = strSQL & "[Last_Thread_ID] [int] NULL,"
			strSQL = strSQL & "[No_of_replies] [int] NOT NULL DEFAULT (0),"
			strSQL = strSQL & "[No_of_views] [int] NOT NULL DEFAULT (0),"
			strSQL = strSQL & "[Priority] [smallint] NOT NULL, "
			strSQL = strSQL & "[Locked] [bit] NOT NULL DEFAULT (0),"
			strSQL = strSQL & "[Hide] [bit] NOT NULL DEFAULT (0),"
			strSQL = strSQL & "[Event_date] [datetime] NULL, "
			strSQL = strSQL & "[Event_date_end] [datetime] NULL, "
			strSQL = strSQL & "[Rating] [float] NOT NULL DEFAULT (0) ,"
			strSQL = strSQL & "[Rating_Total] [int] NOT NULL DEFAULT (0) ,"
			strSQL = strSQL & "[Rating_Votes] [int] NOT NULL DEFAULT (0) "
			strSQL = strSQL & ") ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "Topic <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
	
	
			'Create the Thread Table
			strSQL = "CREATE TABLE [" & strDBO & "].[" & strDbTable & "Thread] ("
			strSQL = strSQL & "[Thread_ID] [int] IDENTITY (1, 1) PRIMARY KEY  CLUSTERED  NOT NULL ,"
			strSQL = strSQL & "[Topic_ID] [int] NOT NULL ,"
			strSQL = strSQL & "[Author_ID] [int] NULL ,"
			strSQL = strSQL & "[Message_date] [datetime] NOT NULL  DEFAULT (getdate()) ,"
			strSQL = strSQL & "[IP_addr] [nvarchar] (50) NULL, "
			strSQL = strSQL & "[Show_signature] [bit] NOT NULL DEFAULT (0), "
			strSQL = strSQL & "[Hide] [bit] NOT NULL DEFAULT (0), "
			strSQL = strSQL & "[Answer] [bit] NOT NULL DEFAULT (0) ,"
			strSQL = strSQL & "[Thanks] [int] NOT NULL DEFAULT (0) ,"
			strSQL = strSQL & "[Message] [text] NULL " 
			strSQL = strSQL & ") ON [PRIMARY]"
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "Thread <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
	
	
	
			'Create the Author Table
			strSQL = "CREATE TABLE [" & strDBO & "].[" & strDbTable & "Author] ("
			strSQL = strSQL & "[Author_ID] [int] IDENTITY (1, 1) PRIMARY KEY  CLUSTERED  NOT NULL ,"
			strSQL = strSQL & "[Group_ID] [int] NOT NULL ,"
			strSQL = strSQL & "[Username] [nvarchar] (35) NOT NULL ,"
			strSQL = strSQL & "[Real_name] [nvarchar] (40) NULL ,"
			strSQL = strSQL & "[User_code] [nvarchar] (50) NOT NULL ,"
			strSQL = strSQL & "[Password] [nvarchar] (50) NOT NULL ,"
			strSQL = strSQL & "[Salt] [nvarchar] (30) NULL ,"
			strSQL = strSQL & "[Author_email] [nvarchar] (75) NULL ,"
			strSQL = strSQL & "[Gender] [nvarchar] (10) NULL ,"
			strSQL = strSQL & "[Photo] [nvarchar] (100) NULL ,"
			strSQL = strSQL & "[Homepage] [nvarchar] (50) NULL ,"
			strSQL = strSQL & "[Location] [nvarchar] (60) NULL ,"
			strSQL = strSQL & "[MSN] [nvarchar] (75) NULL ,"
			strSQL = strSQL & "[Yahoo] [nvarchar] (75) NULL ,"
			strSQL = strSQL & "[ICQ] [nvarchar] (20) NULL ,"
			strSQL = strSQL & "[AIM] [nvarchar] (75) NULL ,"
			strSQL = strSQL & "[Occupation] [nvarchar] (60) NULL ,"
			strSQL = strSQL & "[Interests] [nvarchar] (160) NULL ,"
			strSQL = strSQL & "[DOB] [datetime] NULL ,"
			strSQL = strSQL & "[Signature] [nvarchar] (255) NULL ,"
			strSQL = strSQL & "[No_of_posts] [int] NOT NULL DEFAULT (0),"
			strSQL = strSQL & "[Points] [int] NOT NULL DEFAULT (0) ,"
			strSQL = strSQL & "[Join_date] [datetime] NOT NULL DEFAULT (getdate()),"
			strSQL = strSQL & "[Avatar] [nvarchar] (100) NULL ,"
			strSQL = strSQL & "[Avatar_title] [nvarchar] (70) NULL ,"
			strSQL = strSQL & "[Last_visit] [datetime] NOT NULL DEFAULT (getdate()),"
			strSQL = strSQL & "[Time_offset] [nvarchar] (1) NOT NULL ,"
			strSQL = strSQL & "[Time_offset_hours] [smallint] NOT NULL ,"
			strSQL = strSQL & "[Date_format] [nvarchar] (10) NULL ,"
			strSQL = strSQL & "[No_of_PM] [int] NOT NULL DEFAULT (0),"
			strSQL = strSQL & "[Inbox_no_of_PM] [int] DEFAULT (0), "
			strSQL = strSQL & "[Show_email] [bit] NOT NULL ,"
			strSQL = strSQL & "[Attach_signature] [bit] NOT NULL ,"
			strSQL = strSQL & "[Active] [bit] NOT NULL ,"
			strSQL = strSQL & "[Rich_editor] [bit] NOT NULL ,"
			strSQL = strSQL & "[Reply_notify] [bit] NOT NULL ,"
			strSQL = strSQL & "[PM_notify] [bit] NOT NULL, "
			strSQL = strSQL & "[Skype] [nvarchar] (30) NULL ,"
			strSQL = strSQL & "[Login_attempt] [int] NOT NULL DEFAULT (0), "
			strSQL = strSQL & "[Banned] [bit] NOT NULL DEFAULT (0),"
			strSQL = strSQL & "[Newsletter] [bit] NOT NULL DEFAULT (0),"
			strSQL = strSQL & "[Info] [nvarchar] (255) NULL ,"
			strSQL = strSQL & "[Login_IP] [nvarchar] (50) NULL ,"
			strSQL = strSQL & "[Custom1] [nvarchar] (255) NULL ,"
			strSQL = strSQL & "[Custom2] [nvarchar] (255) NULL ,"
			strSQL = strSQL & "[Custom3] [nvarchar] (255) NULL ,"
			strSQL = strSQL & "[Answered] [int] NOT NULL DEFAULT (0) ,"
			strSQL = strSQL & "[Thanked] [int] NOT NULL DEFAULT (0) ,"
			strSQL = strSQL & "[LinkedIn] [nvarchar] (75) NULL ,"
			strSQL = strSQL & "[Facebook] [nvarchar] (75) NULL ,"
			strSQL = strSQL & "[Twitter] [nvarchar] (75) NULL "
			strSQL = strSQL & ") ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "Author <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
	
	
	
			'Create the Private messages table
			strSQL = "CREATE TABLE [" & strDBO & "].[" & strDbTable & "PMMessage] ("
			strSQL = strSQL & "[PM_ID] [int] IDENTITY (1, 1) PRIMARY KEY  CLUSTERED  NOT NULL ,"
			strSQL = strSQL & "[Author_ID] [int] NOT NULL ,"
			strSQL = strSQL & "[From_ID] [int] NOT NULL ,"
			strSQL = strSQL & "[PM_Tittle] [nvarchar] (70) NOT NULL ,"
			strSQL = strSQL & "[PM_Message] [text] NOT NULL ," 'This is set to a text datatype inorder to hold large posts as nvarchar has 4,000 limit and varchar an 8,000 limit
			strSQL = strSQL & "[PM_Message_date] [datetime] NOT NULL  DEFAULT (getdate()),"
			strSQL = strSQL & "[Read_Post] [bit] NOT NULL DEFAULT (0),"
			strSQL = strSQL & "[Email_notify] [bit] NOT NULL DEFAULT (0), "
			strSQL = strSQL & "[Inbox] [bit] NOT NULL DEFAULT (1), "
			strSQL = strSQL & "[Outbox] [bit] NOT NULL DEFAULT (1) "
			strSQL = strSQL & ") ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "PMMessage <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
	
	
	
			'Create the Buddy List table
			strSQL = "CREATE TABLE [" & strDBO & "].[" & strDbTable & "BuddyList] ("
			strSQL = strSQL & "[Address_ID] [int] IDENTITY (1, 1) PRIMARY KEY  CLUSTERED  NOT NULL ,"
			strSQL = strSQL & "[Author_ID] [int] NOT NULL ,"
			strSQL = strSQL & "[Buddy_ID] [int] NOT NULL ,"
			strSQL = strSQL & "[Description] [nvarchar] (60) NOT NULL ,"
			strSQL = strSQL & "[Block] [bit] NOT NULL DEFAULT (0) "
			strSQL = strSQL & ") ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "BuddyList <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
	
	
	
	
			'Create the Date Time Table
			strSQL = "CREATE TABLE [" & strDBO & "].[" & strDbTable & "DateTimeFormat] ("
			strSQL = strSQL & "[ID] [int] IDENTITY (1, 1) PRIMARY KEY  CLUSTERED  NOT NULL ,"
			strSQL = strSQL & "[Date_format] [nvarchar] (10)  NOT NULL ,"
			strSQL = strSQL & "[Year_format] [nvarchar] (6)  NULL ,"
			strSQL = strSQL & "[Seporator] [nvarchar] (15)  NULL ,"
			strSQL = strSQL & "[Month1] [nvarchar] (15)  NULL ,"
			strSQL = strSQL & "[Month2] [nvarchar] (15)  NULL ,"
			strSQL = strSQL & "[Month3] [nvarchar] (15)  NULL ,"
			strSQL = strSQL & "[Month4] [nvarchar] (15)  NULL ,"
			strSQL = strSQL & "[Month5] [nvarchar] (15)  NULL ,"
			strSQL = strSQL & "[Month6] [nvarchar] (15)  NULL ,"
			strSQL = strSQL & "[Month7] [nvarchar] (15)  NULL ,"
			strSQL = strSQL & "[Month8] [nvarchar] (15)  NULL ,"
			strSQL = strSQL & "[Month9] [nvarchar] (15)  NULL ,"
			strSQL = strSQL & "[Month10] [nvarchar] (15)  NULL ,"
			strSQL = strSQL & "[Month11] [nvarchar] (15)  NULL ,"
			strSQL = strSQL & "[Month12] [nvarchar] (15)  NULL ,"
			strSQL = strSQL & "[Time_format] [smallint] NULL ,"
			strSQL = strSQL & "[am] [nvarchar] (6)  NULL ,"
			strSQL = strSQL & "[pm] [nvarchar] (6)  NULL, "
			strSQL = strSQL & "[Time_offset] [nvarchar] (1) NOT NULL ,"
			strSQL = strSQL & "[Time_offset_hours] [smallint] NOT NULL "
			strSQL = strSQL & ") ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "DateTimeFormat <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			'Create the Group Table
			strSQL = "CREATE TABLE [" & strDBO & "].[" & strDbTable & "Group] ("
			strSQL = strSQL & "[Group_ID] [int] IDENTITY (1, 1)  PRIMARY KEY  CLUSTERED  NOT NULL ,"
			strSQL = strSQL & "[Name] [nvarchar] (40) NULL ,"
			strSQL = strSQL & "[Minimum_posts] [int] NOT NULL DEFAULT (0),"
			strSQL = strSQL & "[Special_rank] [bit] NOT NULL ,"
			strSQL = strSQL & "[Stars] [int] NOT NULL DEFAULT (0),"
			strSQL = strSQL & "[Custom_stars] [nvarchar] (80) NULL, "
			strSQL = strSQL & "[Starting_group] [bit] NOT NULL DEFAULT (0), "
			strSQL = strSQL & "[Image_uploads] [bit] NOT NULL DEFAULT (0), "
			strSQL = strSQL & "[File_uploads] [bit] NOT NULL DEFAULT (0), "
			strSQL = strSQL & "[Ladder_ID] [int] NOT NULL DEFAULT (0) ,"
			strSQL = strSQL & "[Signatures] [bit] NOT NULL DEFAULT (1), "
			strSQL = strSQL & "[URLs] [bit] NOT NULL DEFAULT (1), "
			strSQL = strSQL & "[Images] [bit] NOT NULL DEFAULT (1), "
			strSQL = strSQL & "[Private_Messenger] [bit] NOT NULL DEFAULT (1), "
			strSQL = strSQL & "[Chat_Room] [bit] NOT NULL DEFAULT (1) "
			strSQL = strSQL & ") ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "Group <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			'Create the Poll Table
			strSQL = "CREATE TABLE [" & strDBO & "].[" & strDbTable & "Poll] ("
			strSQL = strSQL & "[Poll_ID] [int] IDENTITY (1, 1)  PRIMARY KEY  CLUSTERED  NOT NULL ,"
			strSQL = strSQL & "[Poll_question] [nvarchar] (90) NULL ,"
			strSQL = strSQL & "[Multiple_votes] [bit] NOT NULL ,"
			strSQL = strSQL & "[Reply] [bit] NOT NULL "
			strSQL = strSQL & ") ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "Poll <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			'Create the Poll Choices Table
			strSQL = "CREATE TABLE [" & strDBO & "].[" & strDbTable & "PollChoice] ("
			strSQL = strSQL & "[Choice_ID] [int] IDENTITY (1, 1)  PRIMARY KEY  CLUSTERED  NOT NULL ,"
			strSQL = strSQL & "[Poll_ID] [int] NOT NULL DEFAULT (0),"
			strSQL = strSQL & "[Choice] [nvarchar] (80) NULL ,"
			strSQL = strSQL & "[Votes] [int] NOT NULL DEFAULT (0)"
			strSQL = strSQL & ") ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write("<br />Error Creating the Table " & strDbTable & "PollChoice <br />" & Err.description & "<br /><br />")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			'Create the Poll Vote Table
			strSQL = "CREATE TABLE [" & strDBO & "].[" & strDbTable & "PollVote] ("
			strSQL = strSQL & "[Poll_ID] [int] NOT NULL DEFAULT (0),"
			strSQL = strSQL & "[Author_ID] [int] NOT NULL DEFAULT (0)"
			strSQL = strSQL & ")"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "PollVote <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			'Create the Email Notify Table
			strSQL = "CREATE TABLE [" & strDBO & "].[" & strDbTable & "EmailNotify] ("
			strSQL = strSQL & "[Watch_ID] [int] IDENTITY (1, 1)  PRIMARY KEY  CLUSTERED  NOT NULL ,"
			strSQL = strSQL & "[Author_ID] [int] NOT NULL DEFAULT (0),"
			strSQL = strSQL & "[Forum_ID] [int] NOT NULL DEFAULT (0),"
			strSQL = strSQL & "[Topic_ID] [int] NOT NULL DEFAULT (0)"
			strSQL = strSQL & ") ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "EmailNotify <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			'Create the Permissions Table
			strSQL = "CREATE TABLE [" & strDBO & "].[" & strDbTable & "Permissions] ("
			strSQL = strSQL & "[Group_ID] [int] NOT NULL DEFAULT (0),"
			strSQL = strSQL & "[Author_ID] [int] NOT NULL DEFAULT (0),"
			strSQL = strSQL & "[Forum_ID] [int] NOT NULL DEFAULT (0),"
			strSQL = strSQL & "[View_Forum] [bit] NOT NULL ,"
			strSQL = strSQL & "[Post] [bit] NOT NULL ,"
			strSQL = strSQL & "[Reply_posts] [bit] NOT NULL ,"
			strSQL = strSQL & "[Edit_posts] [bit] NOT NULL ,"
			strSQL = strSQL & "[Delete_posts] [bit] NOT NULL ,"
			strSQL = strSQL & "[Priority_posts] [bit] NOT NULL ,"
			strSQL = strSQL & "[Poll_create] [bit] NOT NULL ,"
			strSQL = strSQL & "[Vote] [bit] NOT NULL ,"
			strSQL = strSQL & "[Attachments] [bit] NOT NULL ,"
			strSQL = strSQL & "[Image_upload] [bit] NOT NULL ,"
			strSQL = strSQL & "[Moderate] [bit] NOT NULL,"
			strSQL = strSQL & "[Display_post] [bit] NOT NULL,"
			strSQL = strSQL & "[Calendar_event] [bit] NOT NULL"
			strSQL = strSQL & ") ON [PRIMARY]"
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "Permissions <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			
			'Create the Thread Table
			strSQL = "CREATE TABLE [" & strDBO & "].[" & strDbTable & "Session] ("
			strSQL = strSQL & "[Session_ID] [nvarchar] (50) PRIMARY KEY  CLUSTERED  NOT NULL ,"
			strSQL = strSQL & "[IP_address] [nvarchar] (50) NOT NULL, "
			strSQL = strSQL & "[Last_active] [datetime] NOT NULL  DEFAULT (getdate()) ,"
			strSQL = strSQL & "[Session_data] [nvarchar] (255) NULL "
			strSQL = strSQL & ") ON [PRIMARY]"
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "Session <br />" & Err.description  & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
	
	
			'Create the GuestName Table
			strSQL = "CREATE TABLE [" & strDBO & "].[" & strDbTable & "GuestName] ("
			strSQL = strSQL & "[Guest_ID] [int] IDENTITY (1, 1)  PRIMARY KEY  CLUSTERED  NOT NULL ,"
			strSQL = strSQL & "[Thread_ID] [int] NULL ,"
			strSQL = strSQL & "[Name] [nvarchar] (30)  NULL"
			strSQL = strSQL & ") ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "GuestName <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
	
	
	
	
	
			'Create the Smut Table
			strSQL = "CREATE TABLE [" & strDBO & "].[" & strDbTable & "Smut] ("
			strSQL = strSQL & "[ID_no] [int] IDENTITY (1, 1)  PRIMARY KEY  CLUSTERED  NOT NULL ,"
			strSQL = strSQL & "[Smut] [nvarchar] (50)  NULL ,"
			strSQL = strSQL & "[Word_replace] [nvarchar] (50)  NULL"
			strSQL = strSQL & ") ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "Smut <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			'Create the Ban Table
			strSQL = "CREATE TABLE [" & strDBO & "].[" & strDbTable & "BanList] ("
			strSQL = strSQL & "[Ban_ID] [int] IDENTITY (1, 1)  PRIMARY KEY  CLUSTERED  NOT NULL ,"
			strSQL = strSQL & "[IP] [nvarchar] (30) NULL ,"
			strSQL = strSQL & "[Email] [nvarchar] (60) NULL ,"
			strSQL = strSQL & "[Reason] [nvarchar] (50) NULL "
			strSQL = strSQL & ") ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "BanList <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
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
	
	
			
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<strong>Complete</strong>';" & _
			vbCrLf & "</script>")
			
	
	
	
	
	'******************************************
	'***  		 Create indexes	      *****
	'******************************************

			'Stage 2 start
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br /><strong>Stage 2: Creating Database Index\'s..... </strong>';" & _
			vbCrLf & "</script>")
			
	
	
			strSQL = "CREATE  INDEX [Ban_ID] ON [" & strDBO & "].[" & strDbTable & "BanList]([Ban_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  UNIQUE  INDEX [Cat_ID] ON [" & strDBO & "].[" & strDbTable & "Category]([Cat_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [" & strDbTable & "Group_ID] ON [" & strDBO & "].[" & strDbTable & "Group]([Group_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [Poll_ID] ON [" & strDBO & "].[" & strDbTable & "Poll]([Poll_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [Author_ID] ON [" & strDBO & "].[" & strDbTable & "Author]([Author_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [Group_ID] ON [" & strDBO & "].[" & strDbTable & "Author]([Group_ID]) ON [PRIMARY]"
			
			'Write to the database
			adoCon.Execute(strSQL)
			
			
			strSQL = "CREATE  INDEX [" & strDbTable & "Group" & strDbTable & "Author] ON [" & strDBO & "].[" & strDbTable & "Author]([Group_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  UNIQUE  INDEX [User_code] ON [" & strDBO & "].[" & strDbTable & "Author]([User_code]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  UNIQUE  INDEX [Username] ON [" & strDBO & "].[" & strDbTable & "Author]([Username]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [Cat_ID] ON [" & strDBO & "].[" & strDbTable & "Forum]([Cat_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [Sub_ID] ON [" & strDBO & "].[" & strDbTable & "Forum]([Sub_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [Last_post_author_ID] ON [" & strDBO & "].[" & strDbTable & "Forum]([Last_post_author_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			
			strSQL = "CREATE  INDEX [Choice_ID] ON [" & strDBO & "].[" & strDbTable & "PollChoice]([Choice_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [Poll_ID] ON [" & strDBO & "].[" & strDbTable & "PollChoice]([Poll_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [" & strDbTable & "Polls" & strDbTable & "PollChoice] ON [" & strDBO & "].[" & strDbTable & "PollChoice]([Poll_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [Address_ID] ON [" & strDBO & "].[" & strDbTable & "BuddyList]([Address_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [Author_ID] ON [" & strDBO & "].[" & strDbTable & "BuddyList]([Buddy_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [Buddy_ID] ON [" & strDBO & "].[" & strDbTable & "BuddyList]([Author_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [" & strDbTable & "Author" & strDbTable & "BuddyList] ON [" & strDBO & "].[" & strDbTable & "BuddyList]([Buddy_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [Author_ID] ON [" & strDBO & "].[" & strDbTable & "EmailNotify]([Author_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [Forum_ID] ON [" & strDBO & "].[" & strDbTable & "EmailNotify]([Forum_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [" & strDbTable & "Author" & strDbTable & "TopicWatch] ON [" & strDBO & "].[" & strDbTable & "EmailNotify]([Author_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [Toipc_ID] ON [" & strDBO & "].[" & strDbTable & "EmailNotify]([Topic_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [Watch_ID] ON [" & strDBO & "].[" & strDbTable & "EmailNotify]([Watch_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [Auhor_ID] ON [" & strDBO & "].[" & strDbTable & "PMMessage]([Author_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [From_ID] ON [" & strDBO & "].[" & strDbTable & "PMMessage]([From_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [Message_ID] ON [" & strDBO & "].[" & strDbTable & "PMMessage]([PM_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			
			strSQL = "CREATE  INDEX [" & strDbTable & "Forum_ID] ON [" & strDBO & "].[" & strDbTable & "Permissions]([Forum_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			
			strSQL = "CREATE  INDEX [Forum_ID] ON [" & strDBO & "].[" & strDbTable & "Topic]([Forum_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [Poll_ID] ON [" & strDBO & "].[" & strDbTable & "Topic]([Poll_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [Start_Thread_ID] ON [" & strDBO & "].[" & strDbTable & "Topic]([Start_Thread_ID] DESC) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [Last_Thread_ID] ON [" & strDBO & "].[" & strDbTable & "Topic]([Last_Thread_ID] DESC) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [Moved_ID] ON [" & strDBO & "].[" & strDbTable & "Topic]([Moved_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			
			strSQL = "CREATE  INDEX [Topic_ID] ON [" & strDBO & "].[" & strDbTable & "Topic]([Topic_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [Message_date] ON [" & strDBO & "].[" & strDbTable & "Thread]([Message_date] DESC) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [Message_ID] ON [" & strDBO & "].[" & strDbTable & "Thread]([Thread_ID] DESC) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [" & strDbTable & "Author" & strDbTable & "Thread] ON [" & strDBO & "].[" & strDbTable & "Thread]([Author_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [" & strDbTable & "Topic" & strDbTable & "Thread] ON [" & strDBO & "].[" & strDbTable & "Thread]([Topic_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX [Topic_ID] ON [" & strDBO & "].[" & strDbTable & "Thread]([Topic_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
					
			strSQL = "CREATE  INDEX [Guest_ID] ON [" & strDBO & "].[" & strDbTable & "GuestName]([Guest_ID]) ON [PRIMARY] "
			
			'Write to the database
			adoCon.Execute(strSQL)
			
	 		strSQL = "CREATE  INDEX [" & strDbTable & "Thread" & strDbTable & "GuestName] ON [" & strDBO & "].[" & strDbTable & "GuestName]([Thread_ID]) ON [PRIMARY]"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
	 		strSQL = "CREATE  INDEX [Thread_ID] ON [" & strDBO & "].[" & strDbTable & "GuestName]([Thread_ID]) ON [PRIMARY]"
			
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE UNIQUE INDEX [Session_ID] ON [" & strDBO & "].[" & strDbTable & "Session]([Session_ID]) ON [PRIMARY]"
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating one or more Indexs <br />" & Err.description & ".';" & _
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
	'***  	Create relationships	      *****
	'******************************************
	
			'Stage 2 start
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br /><strong>Stage 3: Creating Table Relationships..... </strong>';" & _
			vbCrLf & "</script>")
			
	
			'Create relations between the " & strDbTable & "Author and the " & strDbTable & "Group tables
			strSQL = "ALTER TABLE [" & strDBO & "].[" & strDbTable & "Author] ADD "
			strSQL = strSQL & vbCrLf & "CONSTRAINT [" & strDbTable & "Author_FK00] FOREIGN KEY "
			strSQL = strSQL & vbCrLf & "("
			strSQL = strSQL & vbCrLf & "[Group_ID]"
			strSQL = strSQL & vbCrLf & ") REFERENCES [" & strDBO & "].[" & strDbTable & "Group] ("
			strSQL = strSQL & vbCrLf & "[Group_ID]"
			strSQL = strSQL & vbCrLf & ")"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Relationship between the " & strDbTable & "Author and " & strDbTable & "Group tables <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			'Create relations between the " & strDbTable & "Forum and the " & strDbTable & "Category tables
			strSQL = "ALTER TABLE [" & strDBO & "].[" & strDbTable & "Forum] ADD "
			strSQL = strSQL & vbCrLf & "CONSTRAINT [" & strDbTable & "Forum_FK00] FOREIGN KEY "
			strSQL = strSQL & vbCrLf & "("
			strSQL = strSQL & vbCrLf & "[Cat_ID]"
			strSQL = strSQL & vbCrLf & ") REFERENCES [" & strDBO & "].[" & strDbTable & "Category] ("
			strSQL = strSQL & vbCrLf & "[Cat_ID]"
			strSQL = strSQL & vbCrLf & ")"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Relationship between the " & strDbTable & "Forum and " & strDbTable & "Category tables <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
				
			
			'Create relations between the " & strDbTable & "PollChoice and the " & strDbTable & "Poll tables
			strSQL = "ALTER TABLE [" & strDBO & "].[" & strDbTable & "PollChoice] ADD "
			strSQL = strSQL & vbCrLf & "CONSTRAINT [" & strDbTable & "PollChoice_FK00] FOREIGN KEY "
			strSQL = strSQL & vbCrLf & "("
			strSQL = strSQL & vbCrLf & "[Poll_ID]"
			strSQL = strSQL & vbCrLf & ") REFERENCES [" & strDBO & "].[" & strDbTable & "Poll] ("
			strSQL = strSQL & vbCrLf & "[Poll_ID]"
			strSQL = strSQL & vbCrLf & ")"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Relationship between the " & strDbTable & "PollChoice and " & strDbTable & "Poll tables <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			'Create relations between the " & strDbTable & "BuddyList and the " & strDbTable & "Author tables
			strSQL = "ALTER TABLE [" & strDBO & "].[" & strDbTable & "BuddyList] ADD "
			strSQL = strSQL & vbCrLf & "CONSTRAINT [" & strDbTable & "BuddyList_FK00] FOREIGN KEY "
			strSQL = strSQL & vbCrLf & "("
			strSQL = strSQL & vbCrLf & "[Buddy_ID]"
			strSQL = strSQL & vbCrLf & ") REFERENCES [" & strDBO & "].[" & strDbTable & "Author] ("
			strSQL = strSQL & vbCrLf & "[Author_ID]"
			strSQL = strSQL & vbCrLf & ")"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Relationship between the " & strDbTable & "BuddyList and " & strDbTable & "Author tables <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
	
	
			'Create relations between the " & strDbTable & "EmailNotify and the " & strDbTable & "Author tables
			strSQL = "ALTER TABLE [" & strDBO & "].[" & strDbTable & "EmailNotify] ADD "
			strSQL = strSQL & vbCrLf & "CONSTRAINT [" & strDbTable & "EmailNotify_FK00] FOREIGN KEY "
			strSQL = strSQL & vbCrLf & "("
			strSQL = strSQL & vbCrLf & "[Author_ID]"
			strSQL = strSQL & vbCrLf & ") REFERENCES [" & strDBO & "].[" & strDbTable & "Author] ("
			strSQL = strSQL & vbCrLf & "[Author_ID]"
			strSQL = strSQL & vbCrLf & ")"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Relationship between the " & strDbTable & "EmailNotify and " & strDbTable & "Author tables <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			'Create relations between the PMMessage and the Author tables
			strSQL = "ALTER TABLE [" & strDBO & "].[" & strDbTable & "PMMessage] ADD "
			strSQL = strSQL & vbCrLf & "CONSTRAINT [" & strDbTable & "PMMessage_FK00] FOREIGN KEY "
			strSQL = strSQL & vbCrLf & "("
			strSQL = strSQL & vbCrLf & "[From_ID]"
			strSQL = strSQL & vbCrLf & ") REFERENCES [" & strDBO & "].[" & strDbTable & "Author] ("
			strSQL = strSQL & vbCrLf & "[Author_ID]"
			strSQL = strSQL & vbCrLf & ")"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Relationship between the " & strDbTable & "PMMessage and " & strDbTable & "Author tables <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			'Create relations between the Permissions and the Forum tables
			strSQL = "ALTER TABLE [" & strDBO & "].[" & strDbTable & "Permissions] ADD "
			strSQL = strSQL & vbCrLf & "CONSTRAINT [" & strDbTable & "Permissions_FK00] FOREIGN KEY "
			strSQL = strSQL & vbCrLf & "("
			strSQL = strSQL & vbCrLf & "[Forum_ID]"
			strSQL = strSQL & vbCrLf & ") REFERENCES [" & strDBO & "].[" & strDbTable & "Forum] ("
			strSQL = strSQL & vbCrLf & "[Forum_ID]"
			strSQL = strSQL & vbCrLf & ")"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Relationship between the " & strDbTable & "Permissions and " & strDbTable & "Forum tables <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			'Create relations between the " & strDbTable & "Topic and the " & strDbTable & "Forum tables
			strSQL = "ALTER TABLE [" & strDBO & "].[" & strDbTable & "Topic] ADD "
			strSQL = strSQL & vbCrLf & "CONSTRAINT [" & strDbTable & "Topic_FK00] FOREIGN KEY "
			strSQL = strSQL & vbCrLf & "("
			strSQL = strSQL & vbCrLf & "[Forum_ID]"
			strSQL = strSQL & vbCrLf & ") REFERENCES [" & strDBO & "].[" & strDbTable & "Forum] ("
			strSQL = strSQL & vbCrLf & "[Forum_ID]"
			strSQL = strSQL & vbCrLf & ")"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Relationship between the " & strDbTable & "Topic and " & strDbTable & "Forum tables <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			'Create relations between the GuestName and the Thread tables
			strSQL = "ALTER TABLE [" & strDBO & "].[" & strDbTable & "GuestName] ADD "
			strSQL = strSQL & vbCrLf & "CONSTRAINT [" & strDbTable & "GuestName_FK00] FOREIGN KEY "
			strSQL = strSQL & vbCrLf & "("
			strSQL = strSQL & vbCrLf & "[Thread_ID]"
			strSQL = strSQL & vbCrLf & ") REFERENCES [" & strDBO & "].[" & strDbTable & "Thread] ("
			strSQL = strSQL & vbCrLf & "[Thread_ID]"
			strSQL = strSQL & vbCrLf & ")"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Relationship between the " & strDbTable & "GuestName and " & strDbTable & "Thread tables <br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
	
	
			'Create relations between the " & strDbTable & "Thread, " & strDbTable & "Topic and the " & strDbTable & "Author tables
			strSQL = "ALTER TABLE [" & strDBO & "].[" & strDbTable & "Thread] ADD "
			strSQL = strSQL & vbCrLf & "CONSTRAINT [" & strDbTable & "Thread_FK00] FOREIGN KEY "
			strSQL = strSQL & vbCrLf & "("
			strSQL = strSQL & vbCrLf & "[Author_ID]"
			strSQL = strSQL & vbCrLf & ") REFERENCES [" & strDBO & "].[" & strDbTable & "Author] ("
			strSQL = strSQL & vbCrLf & "[Author_ID]"
			strSQL = strSQL & vbCrLf & "), "
			strSQL = strSQL & vbCrLf & "CONSTRAINT [" & strDbTable & "Thread_FK01] FOREIGN KEY "
			strSQL = strSQL & vbCrLf & "("
			strSQL = strSQL & vbCrLf & "[Topic_ID]"
			strSQL = strSQL & vbCrLf & ") REFERENCES [" & strDBO & "].[" & strDbTable & "Topic] ("
			strSQL = strSQL & vbCrLf & "[Topic_ID]"
			strSQL = strSQL & vbCrLf & ")"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Relationship between the " & strDbTable & "Thread, " & strDbTable & "Topic and " & strDbTable & "Author tables <br />" & Err.description & ".';" & _
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
	
	
			'Stage 4 start
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br /><strong>Stage 4: Entering default values for new fields..... </strong>';" & _
			vbCrLf & "</script>")
			
			
			'Enter the default values in the LadderGroup Table
			'Primary Ladder Group
			strSQL = "INSERT INTO [" & strDBO & "].[" & strDbTable & "LadderGroup] ("
			strSQL = strSQL & "[Ladder_Name] "
			strSQL = strSQL & ") "
			strSQL = strSQL & "VALUES "
			strSQL = strSQL & "('Primary Ladder Group')"
			
			'Write to the database
			adoCon.Execute(strSQL)
			
			
			'Enter the default values in the UserGroup Table
			'Admin Group
			strSQL = "INSERT INTO [" & strDBO & "].[" & strDbTable & "Group] ("
			strSQL = strSQL & "[Name], "
			strSQL = strSQL & "[Minimum_posts], "
			strSQL = strSQL & "[Special_rank], "
			strSQL = strSQL & "[Ladder_ID], "
			strSQL = strSQL & "[Stars], "
			strSQL = strSQL & "[Starting_group] "
			strSQL = strSQL & ") "
			strSQL = strSQL & "VALUES "
			strSQL = strSQL & "('Admin Group', "
			strSQL = strSQL & "'-1', "
			strSQL = strSQL & "'1', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'5', "
			strSQL = strSQL & "'0')"
			
			'Write to the database
			adoCon.Execute(strSQL)
			
			
			'Guest Group
			strSQL = "INSERT INTO [" & strDBO & "].[" & strDbTable & "Group] ("
			strSQL = strSQL & "[Name], "
			strSQL = strSQL & "[Minimum_posts], "
			strSQL = strSQL & "[Special_rank], "
			strSQL = strSQL & "[Ladder_ID], "
			strSQL = strSQL & "[Stars], "
			strSQL = strSQL & "[Starting_group] "
			strSQL = strSQL & ") "
			strSQL = strSQL & "VALUES "
			strSQL = strSQL & "('Guest Group', "
			strSQL = strSQL & "'-1', "
			strSQL = strSQL & "'1', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'0')"
			
			'Write to the database
			adoCon.Execute(strSQL)
			
			
			'Moderator Group
			strSQL = "INSERT INTO [" & strDBO & "].[" & strDbTable & "Group] ("
			strSQL = strSQL & "[Name], "
			strSQL = strSQL & "[Minimum_posts], "
			strSQL = strSQL & "[Special_rank], "
			strSQL = strSQL & "[Ladder_ID], "
			strSQL = strSQL & "[Stars], "
			strSQL = strSQL & "[Starting_group] "
			strSQL = strSQL & ") "
			strSQL = strSQL & "VALUES "
			strSQL = strSQL & "('Moderator Group', "
			strSQL = strSQL & "'-1', "
			strSQL = strSQL & "'1', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'4', "
			strSQL = strSQL & "'0')"
			
			'Write to the database
			adoCon.Execute(strSQL)
			
			
			'Newbie Group
			strSQL = "INSERT INTO [" & strDBO & "].[" & strDbTable & "Group] ("
			strSQL = strSQL & "[Name], "
			strSQL = strSQL & "[Minimum_posts], "
			strSQL = strSQL & "[Special_rank], "
			strSQL = strSQL & "[Ladder_ID], "
			strSQL = strSQL & "[Stars], "
			strSQL = strSQL & "[Starting_group] "
			strSQL = strSQL & ") "
			strSQL = strSQL & "VALUES "
			strSQL = strSQL & "('Newbie', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'1', "
			strSQL = strSQL & "'1', "
			strSQL = strSQL & "'1')"
			
			'Write to the database
			adoCon.Execute(strSQL)
			
			
			'Groupie Group
			strSQL = "INSERT INTO [" & strDBO & "].[" & strDbTable & "Group] ("
			strSQL = strSQL & "[Name], "
			strSQL = strSQL & "[Minimum_posts], "
			strSQL = strSQL & "[Special_rank], "
			strSQL = strSQL & "[Ladder_ID], "
			strSQL = strSQL & "[Stars], "
			strSQL = strSQL & "[Starting_group] "
			strSQL = strSQL & ") "
			strSQL = strSQL & "VALUES "
			strSQL = strSQL & "('Groupie', "
			strSQL = strSQL & "'40', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'1', "
			strSQL = strSQL & "'2', "
			strSQL = strSQL & "'0')"
			
			'Write to the database
			adoCon.Execute(strSQL)
			
			
			'Full Member Group
			strSQL = "INSERT INTO [" & strDBO & "].[" & strDbTable & "Group] ("
			strSQL = strSQL & "[Name], "
			strSQL = strSQL & "[Minimum_posts], "
			strSQL = strSQL & "[Special_rank], "
			strSQL = strSQL & "[Ladder_ID], "
			strSQL = strSQL & "[Stars], "
			strSQL = strSQL & "[Starting_group] "
			strSQL = strSQL & ") "
			strSQL = strSQL & "VALUES "
			strSQL = strSQL & "('Senior Member', "
			strSQL = strSQL & "'100', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'1', "
			strSQL = strSQL & "'3', "
			strSQL = strSQL & "'0')"
			
			'Write to the database
			adoCon.Execute(strSQL)
			
				
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error entering default values in the Table " & strDbTable & "Group<br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			'Enter the default values in the Author Table
			'Enter the admin account into db
			strSQL = "INSERT INTO [" & strDBO & "].[" & strDbTable & "Author] ("
			strSQL = strSQL & "[Group_ID], "
			strSQL = strSQL & "[Username], "
			strSQL = strSQL & "[User_code], "
			strSQL = strSQL & "[Password], "
			strSQL = strSQL & "[Salt], "
			strSQL = strSQL & "[Show_email], "
			strSQL = strSQL & "[Attach_signature], "
			strSQL = strSQL & "[Time_offset], "
			strSQL = strSQL & "[Time_offset_hours], "
			strSQL = strSQL & "[Rich_editor], "
			strSQL = strSQL & "[Date_format], "
			strSQL = strSQL & "[Active], "
			strSQL = strSQL & "[Reply_notify], "
			strSQL = strSQL & "[PM_notify], "
			strSQL = strSQL & "[No_of_posts], "
			strSQL = strSQL & "[Login_attempt], "
			strSQL = strSQL & "[Banned] "
			strSQL = strSQL & ") "
			strSQL = strSQL & "VALUES "
			strSQL = strSQL & "('1', "
			strSQL = strSQL & "'administrator', "
			strSQL = strSQL & "'administrator2FC73499BAC5A41', "
			strSQL = strSQL & "'A85B3E67CFA695D711570FB9822C0CF82871903B', "
			strSQL = strSQL & "'72964E7', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'+', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'1', "
			strSQL = strSQL & "'dd/mm/yy', "
			strSQL = strSQL & "'1', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'0',"
			strSQL = strSQL & "'0')"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'Enter the Guest account into db
			strSQL = "INSERT INTO [" & strDBO & "].[" & strDbTable & "Author] ("
			strSQL = strSQL & "[Group_ID], "
			strSQL = strSQL & "[Username], "
			strSQL = strSQL & "[User_code], "
			strSQL = strSQL & "[Password], "
			strSQL = strSQL & "[Salt], "
			strSQL = strSQL & "[Show_email], "
			strSQL = strSQL & "[Attach_signature], "
			strSQL = strSQL & "[Time_offset], "
			strSQL = strSQL & "[Time_offset_hours], "
			strSQL = strSQL & "[Rich_editor], "
			strSQL = strSQL & "[Date_format], "
			strSQL = strSQL & "[Active], "
			strSQL = strSQL & "[Reply_notify], "
			strSQL = strSQL & "[PM_notify], "
			strSQL = strSQL & "[No_of_posts], "
			strSQL = strSQL & "[Login_attempt], "
			strSQL = strSQL & "[Banned] "
			strSQL = strSQL & ") "
			strSQL = strSQL & "VALUES "
			strSQL = strSQL & "('2', "
			strSQL = strSQL & "'Guests', "
			strSQL = strSQL & "'Guest48CEE9Z2849AE95A6', "
			strSQL = strSQL & "'6734DDD7A6A6C9F4D34945B0C9CF9677F3221EC9', "
			strSQL = strSQL & "'E4AC', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'+', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'1', "
			strSQL = strSQL & "'dd/mm/yy', "
			strSQL = strSQL & "'1', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'0',"
			strSQL = strSQL & "'0')"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error entering default values in the Table " & strDbTable & "Author<br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
	
	
	
			'Enter the default values in the date time table
			strSQL = "INSERT INTO [" & strDBO & "].[" & strDbTable & "DateTimeFormat] ("
			strSQL = strSQL & "[Date_format], "
			strSQL = strSQL & "[Year_format], "
			strSQL = strSQL & "[Seporator], "
			strSQL = strSQL & "[Month1], "
			strSQL = strSQL & "[Month2], "
			strSQL = strSQL & "[Month3], "
			strSQL = strSQL & "[Month4], "
			strSQL = strSQL & "[Month5], "
			strSQL = strSQL & "[Month6], "
			strSQL = strSQL & "[Month7], "
			strSQL = strSQL & "[Month8], "
			strSQL = strSQL & "[Month9], "
			strSQL = strSQL & "[Month10], "
			strSQL = strSQL & "[Month11], "
			strSQL = strSQL & "[Month12], "
			strSQL = strSQL & "[Time_format], "
			strSQL = strSQL & "[am], "
			strSQL = strSQL & "[pm], "
			strSQL = strSQL & "[Time_offset], "
			strSQL = strSQL & "[Time_offset_hours] "
			strSQL = strSQL & ") "
			strSQL = strSQL & "VALUES "
			strSQL = strSQL & "('dd/mm/yy', "
			strSQL = strSQL & "'long', "
			strSQL = strSQL & "'" & chr(32) & "', "
			strSQL = strSQL & "'Jan', "
			strSQL = strSQL & "'Feb', "
			strSQL = strSQL & "'Mar', "
			strSQL = strSQL & "'Apr', "
			strSQL = strSQL & "'May', "
			strSQL = strSQL & "'Jun', "
			strSQL = strSQL & "'Jul', "
			strSQL = strSQL & "'Aug', "
			strSQL = strSQL & "'Sep', "
			strSQL = strSQL & "'Oct', "
			strSQL = strSQL & "'Nov', "
			strSQL = strSQL & "'Dec', "
			strSQL = strSQL & "'12', "
			strSQL = strSQL & "'am', "
			strSQL = strSQL & "'pm', "
			strSQL = strSQL & "'+', "
			strSQL = strSQL & "'0')"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error entering default values in the Table " & strDbTable & "DateTimeFormat<br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
	
	
	
			'Enter the default values in the smut table
			For intBadWordLoopCounter = 1 to 13
	
				'Write the SQL
				strSQL = "INSERT INTO [" & strDBO & "].[" & strDbTable & "Smut] ([Smut], [Word_replace]) "
				strSQL = strSQL & "VALUES ("
	
				Select Case intBadWordLoopCounter
					Case 1
						strSQL = strSQL & "'\bcunt\b', '<img src=""smileys/smiley35.gif"" border=""0"">'"
					Case 2
						strSQL = strSQL & "'cunting', '<img src=""smileys/smiley35.gif"" border=""0"">'"
					Case 3
						strSQL = strSQL & "'\bfuck\b', '<img src=""smileys/smiley35.gif"" border=""0"">'"
					Case 4
						strSQL = strSQL & "'fucker', '<img src=""smileys/smiley35.gif"" border=""0"">'"
					Case 5
						strSQL = strSQL & "'fucking', '<img src=""smileys/smiley35.gif"" border=""0"">'"
					Case 6
						strSQL = strSQL & "'fuck-off', 'please leave'"
					Case 7
						strSQL = strSQL & "'fuckoff', 'please leave'"
					Case 8
						strSQL = strSQL & "'motherfucker', 'motherf**k'"
					Case 9
						strSQL = strSQL & "'shit', 'sh*t'"
					Case 10
						strSQL = strSQL & "'shiting', 'sh*ting'"
					Case 11
						strSQL = strSQL & "'\bslag\b', 'sl*g'"
					Case 12
						strSQL = strSQL & "'tosser', 't**ser'"
					Case 13
						strSQL = strSQL & "'wanker', 'w**ker'"
				End Select
	
				strSQL = strSQL & ")"
	
				'Write to database
				adoCon.Execute(strSQL)
			Next
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error entering default values in the Table " & strDbTable & "Smut<br />" & Err.description & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			'Intialise the main ADO recordset object
			Set rsCommon = Server.CreateObject("ADODB.Recordset")
			
			
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
			Call addConfigurationItem("Most_active_date", "2017-01-01 00:00:00")
			Call addConfigurationItem("Most_active_users", "1")
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
			
			Set rsCommon = Nothing
			
			
			
			
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<strong>Complete</strong>';" & _
			vbCrLf & "</script>")
			
			
			
	
			'Display a message to say the database is created
			If blnErrorOccured = True Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />" & Err.description & "<br /><br /><h2>Microsoft SQL Server database is set up, but with Error!</h2>'" & _
				vbCrLf & "</script>")
			Else
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br /><h2>Congratulations, Web Wiz Forums Database setup is now complete</h2>'" & _
				vbCrLf & "</script>")
			End If
			
			
			
			'Display completed message
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />The default administrator login for your forum is:-<br /><blockquote>User: Administrator<br />Pass: letmein<br /></blockquote>Click here to go to your <a href=""default.asp"">Forum Homepage</a><br />Click here to login to your <a href=""admin.asp"">Forum Admin Area</a>'" & _
     			vbCrLf & "</script>")
			
			
		End If
	End If

	'Reset Server Variables
	Set adoCon = Nothing
End If
%>