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
If strDatabaseType = "mySQL" AND strSQLDBUserName <> "" Then
	
	

	'Open the database
	Call openDatabase(strCon)

	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then
		
		
		
		Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
		vbCrLf & "	document.getElementById('displayState').innerHTML = '<img src=""forum_images/error.png"" alt=""Error"" /> <strong>Error</strong><br /><strong>Error Connecting to database on mySQL Server</strong><br /><br />Replace the database/database_settings.asp file with the one from the orginal Web Wiz Forums download and start the setup process again.<br /><br /><strong>Error Details:</strong><br />" & Replace(Err.description, "'", "\'") & "';" & _
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
			strSQL = "CREATE TABLE " & strDbTable & "Category ("
			strSQL = strSQL & "Cat_ID INT NOT NULL auto_increment, "
			strSQL = strSQL & "Cat_name varchar(60) NOT NULL , "
			strSQL = strSQL & "Cat_order smallint(4) NOT NULL DEFAULT '1' , "
			strSQL = strSQL & "PRIMARY KEY (Cat_ID));"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
	
				'Write an error message
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "Category <br />" & Replace(Err.description, "'", "\'") & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
	
	
			'Create the Forum Table
			strSQL = "CREATE TABLE " & strDbTable & "Forum ("
			strSQL = strSQL & "Forum_ID INT NOT NULL auto_increment, "
			strSQL = strSQL & "Cat_ID INT NULL ,"
			strSQL = strSQL & "Sub_ID INT NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Forum_Order smallint(4) NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Forum_name varchar(70) NULL ,"
			strSQL = strSQL & "Forum_description varchar(200) NULL ,"
			strSQL = strSQL & "No_of_topics INT NOT NULL DEFAULT '0',"
			strSQL = strSQL & "No_of_posts INT NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Last_post_author_ID INT NOT NULL DEFAULT '1',"
			strSQL = strSQL & "Last_post_date datetime NULL, "
			strSQL = strSQL & "Last_topic_ID  INT NOT NULL DEFAULT '0', "
			strSQL = strSQL & "Password varchar(50) NULL ,"
			strSQL = strSQL & "Forum_code varchar(40) NULL ,"
			strSQL = strSQL & "Show_topics smallint(4) NULL, "
			strSQL = strSQL & "Locked tinyint(1) NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Hide tinyint(1) NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Forum_URL varchar(80) NULL ,"
			strSQL = strSQL & "Forum_icon varchar(80) NULL, "
			strSQL = strSQL & "PRIMARY KEY (Forum_ID));"
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				
				
				'Write an error message
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "Forum<br />" & Replace(Err.description, "'", "\'") & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
	
	
			'Create the Topic Table
			strSQL = "CREATE TABLE " & strDbTable & "Topic ("
			strSQL = strSQL & "Topic_ID INT NOT NULL auto_increment, "
			strSQL = strSQL & "Forum_ID INT NULL ,"
			strSQL = strSQL & "Poll_ID INT NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Moved_ID INT NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Subject varchar(70) NOT NULL,"
			strSQL = strSQL & "Icon varchar(20) NULL,"
			strSQL = strSQL & "Start_Thread_ID INT NULL,"
			strSQL = strSQL & "Last_Thread_ID INT NULL,"
			strSQL = strSQL & "No_of_replies INT NOT NULL DEFAULT '0',"
			strSQL = strSQL & "No_of_views INT NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Priority smallint(4) NOT NULL, "
			strSQL = strSQL & "Locked tinyint(1) NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Hide tinyint(1) NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Event_date datetime NULL,"
			strSQL = strSQL & "Event_date_end datetime NULL,"
			strSQL = strSQL & "Rating DOUBLE NOT NULL DEFAULT '0' ,"
			strSQL = strSQL & "Rating_Total INT NOT NULL DEFAULT '0' ,"
			strSQL = strSQL & "Rating_Votes INT NOT NULL DEFAULT '0', "
			strSQL = strSQL & "PRIMARY KEY (Topic_ID));"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "Topic <br />" & Replace(Err.description, "'", "\'") & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
	
	
			'Create the Thread Table
			strSQL = "CREATE TABLE " & strDbTable & "Thread ("
			strSQL = strSQL & "Thread_ID INT NOT NULL auto_increment, "
			strSQL = strSQL & "Topic_ID INT NOT NULL ,"
			strSQL = strSQL & "Author_ID INT NULL ,"
			strSQL = strSQL & "Message_date datetime NOT NULL ,"
			strSQL = strSQL & "IP_addr varchar(50) NULL, "
			strSQL = strSQL & "Show_signature tinyint(1) NOT NULL DEFAULT '0', "
			strSQL = strSQL & "Hide tinyint(1) NOT NULL DEFAULT '0', "
			strSQL = strSQL & "Answer tinyint(1) NOT NULL DEFAULT '0' ,"
			strSQL = strSQL & "Thanks INT NOT NULL DEFAULT '0', "
			strSQL = strSQL & "Message text NULL ,"
			strSQL = strSQL & "PRIMARY KEY (Thread_ID));"
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "Thread <br />" & Replace(Err.description, "'", "\'") & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
	
	
	
			'Create the Author Table
			strSQL = "CREATE TABLE " & strDbTable & "Author ("
			strSQL = strSQL & "Author_ID INT NOT NULL auto_increment, "
			strSQL = strSQL & "Group_ID INT NOT NULL ,"
			strSQL = strSQL & "Username varchar(35) NOT NULL ,"
			strSQL = strSQL & "Real_name varchar(40) NULL ,"
			strSQL = strSQL & "User_code varchar(50) NOT NULL ,"
			strSQL = strSQL & "Password varchar(50) NOT NULL ,"
			strSQL = strSQL & "Salt varchar(30) NULL ,"
			strSQL = strSQL & "Author_email varchar(75) NULL ,"
			strSQL = strSQL & "Gender varchar(10) NULL ,"
			strSQL = strSQL & "Photo varchar(100) NULL ,"
			strSQL = strSQL & "Homepage varchar(50) NULL ,"
			strSQL = strSQL & "Location varchar(60) NULL ,"
			strSQL = strSQL & "MSN varchar(75) NULL ,"
			strSQL = strSQL & "Yahoo varchar(75) NULL ,"
			strSQL = strSQL & "ICQ varchar(20) NULL ,"
			strSQL = strSQL & "AIM varchar(75) NULL ,"
			strSQL = strSQL & "Occupation varchar(60) NULL ,"
			strSQL = strSQL & "Interests varchar(160) NULL ,"
			strSQL = strSQL & "DOB datetime NULL ,"
			strSQL = strSQL & "Signature varchar(255) NOT NULL DEFAULT '',"
			strSQL = strSQL & "No_of_posts INT NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Points INT NOT NULL DEFAULT '0' ,"
			strSQL = strSQL & "Join_date datetime NOT NULL , "
			strSQL = strSQL & "Avatar varchar(100) NULL ,"
			strSQL = strSQL & "Avatar_title varchar(70) NULL ,"
			strSQL = strSQL & "Last_visit datetime NOT NULL , "
			strSQL = strSQL & "Time_offset varchar(1) NOT NULL ,"
			strSQL = strSQL & "Time_offset_hours smallint(4) NOT NULL ,"
			strSQL = strSQL & "Date_format varchar(10) NULL ,"
			strSQL = strSQL & "No_of_PM INT NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Inbox_no_of_PM INT NOT NULL DEFAULT '0', "
			strSQL = strSQL & "Show_email tinyint(1) NOT NULL ,"
			strSQL = strSQL & "Attach_signature tinyint(1) NOT NULL ,"
			strSQL = strSQL & "Active tinyint(1) NOT NULL ,"
			strSQL = strSQL & "Rich_editor tinyint(1) NOT NULL ,"
			strSQL = strSQL & "Reply_notify tinyint(1) NOT NULL ,"
			strSQL = strSQL & "PM_notify tinyint(1) NOT NULL, "
			strSQL = strSQL & "Skype varchar(30) NULL ,"
			strSQL = strSQL & "Login_attempt INT NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Banned tinyint(1) NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Newsletter tinyint(1) NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Info varchar(255) NOT NULL ,"
			strSQL = strSQL & "Login_IP varchar(50) NULL ,"
			strSQL = strSQL & "Custom1 varchar(255) NULL ,"
			strSQL = strSQL & "Custom2 varchar(255) NULL ,"
			strSQL = strSQL & "Custom3 varchar(255) NULL ,"
			strSQL = strSQL & "Answered INT NOT NULL DEFAULT '0' ,"
			strSQL = strSQL & "Thanked INT NOT NULL DEFAULT '0' ,"
			strSQL = strSQL & "LinkedIn varchar(75) NULL ,"
			strSQL = strSQL & "Facebook varchar(75) NULL ,"
			strSQL = strSQL & "Twitter varchar(75) NULL ,"
			strSQL = strSQL & "PRIMARY KEY (Author_ID));"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
	
	
	
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "Author <br />" & Replace(Err.description, "'", "\'") & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
	
	
	
			'Create the Private messages table
			strSQL = "CREATE TABLE " & strDbTable & "PMMessage ("
			strSQL = strSQL & "PM_ID MEDIUMINT NOT NULL auto_increment, "
			strSQL = strSQL & "Author_ID INT NOT NULL ,"
			strSQL = strSQL & "From_ID INT NOT NULL ,"
			strSQL = strSQL & "PM_Tittle varchar(70) NOT NULL ,"
			strSQL = strSQL & "PM_Message text NOT NULL ," 'This is set to a text datatype inorder to hold large posts as varchar has 4,000 limit and varchar an 8,000 limit
			strSQL = strSQL & "PM_Message_date datetime NOT NULL , "
			strSQL = strSQL & "Read_Post tinyint(1) NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Email_notify tinyint(1) NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Inbox tinyint(1) NOT NULL DEFAULT '1', "
			strSQL = strSQL & "Outbox tinyint(1) NOT NULL DEFAULT '1', "
			strSQL = strSQL & "PRIMARY KEY (PM_ID));"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "PMMessage <br />" & Replace(Err.description, "'", "\'") & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
	
	
	
			'Create the Buddy List table
			strSQL = "CREATE TABLE " & strDbTable & "BuddyList ("
			strSQL = strSQL & "Address_ID INT NOT NULL auto_increment, "
			strSQL = strSQL & "Author_ID INT NOT NULL ,"
			strSQL = strSQL & "Buddy_ID INT NOT NULL ,"
			strSQL = strSQL & "Description varchar(60) NOT NULL ,"
			strSQL = strSQL & "Block tinyint(1) NOT NULL DEFAULT '0', "
			strSQL = strSQL & "PRIMARY KEY (Address_ID));"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "BuddyList <br />" & Replace(Err.description, "'", "\'") & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
	
	
	
	
	
			'Create the Date Time Table
			strSQL = "CREATE TABLE " & strDbTable & "DateTimeFormat ("
			strSQL = strSQL & "ID smallint(4) NOT NULL auto_increment, "
			strSQL = strSQL & "Date_format varchar(10)  NOT NULL  ,"
			strSQL = strSQL & "Year_format varchar(6)  NULL ,"
			strSQL = strSQL & "Seporator varchar(15)  NULL ,"
			strSQL = strSQL & "Month1 varchar(15)  NULL ,"
			strSQL = strSQL & "Month2 varchar(15)  NULL ,"
			strSQL = strSQL & "Month3 varchar(15)  NULL ,"
			strSQL = strSQL & "Month4 varchar(15)  NULL ,"
			strSQL = strSQL & "Month5 varchar(15)  NULL ,"
			strSQL = strSQL & "Month6 varchar(15)  NULL ,"
			strSQL = strSQL & "Month7 varchar(15)  NULL ,"
			strSQL = strSQL & "Month8 varchar(15)  NULL ,"
			strSQL = strSQL & "Month9 varchar(15)  NULL ,"
			strSQL = strSQL & "Month10 varchar(15)  NULL ,"
			strSQL = strSQL & "Month11 varchar(15)  NULL ,"
			strSQL = strSQL & "Month12 varchar(15)  NULL ,"
			strSQL = strSQL & "Time_format smallint(4) NULL ,"
			strSQL = strSQL & "am varchar(6)  NULL ,"
			strSQL = strSQL & "pm varchar(6)  NULL, "
			strSQL = strSQL & "Time_offset varchar(1) NOT NULL ,"
			strSQL = strSQL & "Time_offset_hours smallint(4) NOT NULL ,"
			strSQL = strSQL & "PRIMARY KEY (ID));"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "DateTimeFormat <br />" & Replace(Err.description, "'", "\'") & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			'Create the Group Table
			strSQL = "CREATE TABLE " & strDbTable & "Group ("
			strSQL = strSQL & "Group_ID INT NOT NULL auto_increment, "
			strSQL = strSQL & "Name varchar(40) NULL ,"
			strSQL = strSQL & "Minimum_posts int(4) NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Special_rank tinyint(1) NOT NULL ,"
			strSQL = strSQL & "Stars int(4) NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Custom_stars varchar(80) NULL ,"
			strSQL = strSQL & "Starting_group tinyint(1) NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Image_uploads tinyint(1) NOT NULL DEFAULT '0',"
			strSQL = strSQL & "File_uploads tinyint(1) NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Ladder_ID INT NOT NULL DEFAULT '0' ,"
			strSQL = strSQL & "Signatures tinyint(1) NOT NULL DEFAULT '1', "
			strSQL = strSQL & "URLs tinyint(1) NOT NULL DEFAULT '1', "
			strSQL = strSQL & "Images tinyint(1) NOT NULL DEFAULT '1', "
			strSQL = strSQL & "Private_Messenger tinyint(1) NOT NULL DEFAULT '1', "
			strSQL = strSQL & "Chat_Room tinyint(1) NOT NULL DEFAULT '1', "
			strSQL = strSQL & "PRIMARY KEY (Group_ID));"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "Group <br />" & Replace(Err.description, "'", "\'") & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			'Create the Poll Table
			strSQL = "CREATE TABLE " & strDbTable & "Poll ("
			strSQL = strSQL & "Poll_ID INT NOT NULL auto_increment, "
			strSQL = strSQL & "Poll_question varchar(90) NOT NULL ,"
			strSQL = strSQL & "Multiple_votes tinyint(1) NOT NULL ,"
			strSQL = strSQL & "Reply tinyint(1) NOT NULL ,"
			strSQL = strSQL & "PRIMARY KEY (Poll_ID));"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "Poll <br />" & Replace(Err.description, "'", "\'") & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			'Create the Poll Choices Table
			strSQL = "CREATE TABLE " & strDbTable & "PollChoice ("
			strSQL = strSQL & "Choice_ID INT NOT NULL auto_increment, "
			strSQL = strSQL & "Poll_ID INT NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Choice varchar(80) NOT NULL ,"
			strSQL = strSQL & "Votes INT NOT NULL DEFAULT '0',"
			strSQL = strSQL & "PRIMARY KEY (Choice_ID));"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "PollChoice <br />" & Replace(Err.description, "'", "\'") & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			'Create the Poll Votes Table
			strSQL = "CREATE TABLE " & strDbTable & "PollVote ("
			strSQL = strSQL & "Poll_ID INT NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Author_ID INT NOT NULL DEFAULT '0');"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "PollVote <br />" & Replace(Err.description, "'", "\'") & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			'Create the Email Notify Table
			strSQL = "CREATE TABLE " & strDbTable & "EmailNotify ("
			strSQL = strSQL & "Watch_ID INT NOT NULL auto_increment, "
			strSQL = strSQL & "Author_ID INT NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Forum_ID INT NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Topic_ID INT NOT NULL DEFAULT '0',"
			strSQL = strSQL & "PRIMARY KEY (Watch_ID));"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "EmailNotify <br />" & Replace(Err.description, "'", "\'") & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			'Create the Permissions Table
			strSQL = "CREATE TABLE " & strDbTable & "Permissions ("
			strSQL = strSQL & "Group_ID INT NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Author_ID INT NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Forum_ID INT NOT NULL DEFAULT '0',"
			strSQL = strSQL & "View_Forum tinyint(1) NOT NULL ,"
			strSQL = strSQL & "Post tinyint(1) NOT NULL ,"
			strSQL = strSQL & "Reply_posts tinyint(1) NOT NULL ,"
			strSQL = strSQL & "Edit_posts tinyint(1) NOT NULL ,"
			strSQL = strSQL & "Delete_posts tinyint(1) NOT NULL ,"
			strSQL = strSQL & "Priority_posts tinyint(1) NOT NULL ,"
			strSQL = strSQL & "Poll_create tinyint(1) NOT NULL ,"
			strSQL = strSQL & "Vote tinyint(1) NOT NULL ,"
			strSQL = strSQL & "Attachments tinyint(1) NOT NULL ,"
			strSQL = strSQL & "Image_upload tinyint(1) NOT NULL ,"
			strSQL = strSQL & "Moderate tinyint(1) NOT NULL,"
			strSQL = strSQL & "Display_post tinyint(1) NOT NULL,"
			strSQL = strSQL & "Calendar_event tinyint(1) NOT NULL"
			strSQL = strSQL & ");"
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "Permissions <br />" & Replace(Err.description, "'", "\'") & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			
			'Create the Thread Table
			strSQL = "CREATE TABLE " & strDbTable & "Session ("
			strSQL = strSQL & "Session_ID varchar(50) NOT NULL, "
			strSQL = strSQL & "IP_address varchar(50) NOT NULL, "
			strSQL = strSQL & "Last_active datetime NOT NULL ,"
			strSQL = strSQL & "Session_data varchar (255) NULL ,"
			strSQL = strSQL & "PRIMARY KEY (Session_ID));"
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "Session <br />" & Replace(Err.description, "'", "\'") & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
	
	
			'Create the GuestName Table
			strSQL = "CREATE TABLE " & strDbTable & "GuestName ("
			strSQL = strSQL & "Guest_ID INT NOT NULL auto_increment, "
			strSQL = strSQL & "Thread_ID INT NULL ,"
			strSQL = strSQL & "Name varchar(30)  NULL,"
			strSQL = strSQL & "PRIMARY KEY (Guest_ID));"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "GuestName <br />" & Replace(Err.description, "'", "\'") & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
	
	
	
	
	
			'Create the Smut Table
			strSQL = "CREATE TABLE " & strDbTable & "Smut ("
			strSQL = strSQL & "ID_no smallint(4) NOT NULL auto_increment, "
			strSQL = strSQL & "Smut varchar(50)  NULL ,"
			strSQL = strSQL & "Word_replace varchar(50)  NULL,"
			strSQL = strSQL & "PRIMARY KEY (ID_no));"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "Smut <br />" & Replace(Err.description, "'", "\'") & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			'Create the Ban Table
			strSQL = "CREATE TABLE " & strDbTable & "BanList ("
			strSQL = strSQL & "Ban_ID INT NOT NULL auto_increment, "
			strSQL = strSQL & "IP varchar(30) NULL ,"
			strSQL = strSQL & "Email varchar(60) NULL,"
			strSQL = strSQL & "Reason varchar(50) NULL, "
			strSQL = strSQL & "PRIMARY KEY (Ban_ID));"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating the Table " & strDbTable & "BanList <br />" & Replace(Err.description, "'", "\'") & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			'Create the SetupOptions Table
			strSQL = "CREATE TABLE " & strDbTable & "SetupOptions ("
			strSQL = strSQL & "Option_Item varchar(30),"
			strSQL = strSQL & "Option_Value text NULL " 
			strSQL = strSQL & ")"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br />Error Creating the Table " & strDbTable & "SetupOptions <br />" & Replace(Err.description, "'", "\'") & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			'Create the LadderGroup Table
			strSQL = "CREATE TABLE " & strDbTable & "LadderGroup ("
			strSQL = strSQL & "Ladder_ID INT NOT NULL auto_increment, "
			strSQL = strSQL & "Ladder_Name varchar(25) ,"
			strSQL = strSQL & "PRIMARY KEY (Ladder_ID));"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br />Error Creating the Table " & strDbTable & "LadderGroup <br />" & Replace(Err.description, "'", "\'") & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			'Create the Spam Table
			strSQL = "CREATE TABLE " & strDbTable & "Spam ("
			strSQL = strSQL & "Spam_ID INT NOT NULL auto_increment, "
			strSQL = strSQL & "Spam varchar(255), " 
			strSQL = strSQL & "Spam_Action varchar(20), "
			strSQL = strSQL & "PRIMARY KEY (Spam_ID));"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br />Error Creating the Table " & strDbTable & "Spam <br />" & Replace(Err.description, "'", "\'") & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			'Create the TopicRatingVote Table
			strSQL = "CREATE TABLE " & strDbTable & "TopicRatingVote ("
			strSQL = strSQL & "Topic_ID INT NOT NULL DEFAULT '0', "
			strSQL = strSQL & "Author_ID INT NOT NULL DEFAULT '0' "
			strSQL = strSQL & ")"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br />Error Creating the Table " & strDbTable & "TopicRatingVote <br />" & Replace(Err.description, "'", "\'") & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			'Create the ThreadThanks Table
			strSQL = "CREATE TABLE " & strDbTable & "ThreadThanks ("
			strSQL = strSQL & "Thread_ID INT NOT NULL DEFAULT '0',"
			strSQL = strSQL & "Author_ID INT NOT NULL DEFAULT '0' " 
			strSQL = strSQL & ")"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br />Error Creating the Table " & strDbTable & "ThreadThanks <br />" & Replace(Err.description, "'", "\'") & ".';" & _
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
			
	
	
			strSQL = "CREATE  INDEX Ban_ID ON " & strDbTable & "BanList(Ban_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  UNIQUE  INDEX Cat_ID ON " & strDbTable & "Category(Cat_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX " & strDbTable & "Group_ID ON " & strDbTable & "Group(Group_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE  INDEX Poll_ID ON " & strDbTable & "Poll(Poll_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX Author_ID ON " & strDbTable & "Author(Author_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX Group_ID ON " & strDbTable & "Author(Group_ID);"
			
			'Write to the database
			adoCon.Execute(strSQL)
			
			
			strSQL = "CREATE INDEX " & strDbTable & "Group" & strDbTable & "Author ON " & strDbTable & "Author(Group_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE UNIQUE INDEX User_code ON " & strDbTable & "Author(User_code);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE UNIQUE INDEX Username ON " & strDbTable & "Author(Username);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX Cat_ID ON " & strDbTable & "Forum(Cat_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX Sub_ID ON " & strDbTable & "Forum(Sub_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX Last_post_author_ID ON " & strDbTable & "Forum(Last_post_author_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
					
			strSQL = "CREATE INDEX " & strDbTable & "Categories" & strDbTable & "Forum ON " & strDbTable & "Forum(Cat_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX Choice_ID ON " & strDbTable & "PollChoice(Choice_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX Poll_ID ON " & strDbTable & "PollChoice(Poll_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX " & strDbTable & "Polls" & strDbTable & "PollChoice ON " & strDbTable & "PollChoice(Poll_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX Address_ID ON " & strDbTable & "BuddyList(Address_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX Author_ID ON " & strDbTable & "BuddyList(Buddy_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX Buddy_ID ON " & strDbTable & "BuddyList(Author_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX " & strDbTable & "Author" & strDbTable & "BuddyList ON " & strDbTable & "BuddyList(Buddy_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX Author_ID ON " & strDbTable & "EmailNotify(Author_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX Forum_ID ON " & strDbTable & "EmailNotify(Forum_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX " & strDbTable & "Author" & strDbTable & "TopicWatch ON " & strDbTable & "EmailNotify(Author_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX Toipc_ID ON " & strDbTable & "EmailNotify(Topic_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX Watch_ID ON " & strDbTable & "EmailNotify(Watch_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX Auhor_ID ON " & strDbTable & "PMMessage(Author_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX From_ID ON " & strDbTable & "PMMessage(From_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX Message_ID ON " & strDbTable & "PMMessage(PM_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX " & strDbTable & "Author" & strDbTable & "PMMessage ON " & strDbTable & "PMMessage(From_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX " & strDbTable & "Forum_ID ON " & strDbTable & "Permissions(Forum_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX " & strDbTable & "Forum" & strDbTable & "Permissions ON " & strDbTable & "Permissions(Forum_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX " & strDbTable & "Group_ID ON " & strDbTable & "Permissions(Group_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX Forum_ID ON " & strDbTable & "Topic(Forum_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX Poll_ID ON " & strDbTable & "Topic(Poll_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX Start_Thread_ID ON " & strDbTable & "Topic(Start_Thread_ID DESC);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX Last_Thread_ID ON " & strDbTable & "Topic(Last_Thread_ID DESC);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			
			strSQL = "CREATE INDEX " & strDbTable & "Forum" & strDbTable & "Topic ON " & strDbTable & "Topic(Forum_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX Topic_ID ON " & strDbTable & "Topic(Topic_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX Message_date ON " & strDbTable & "Thread(Message_date DESC);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX Message_ID ON " & strDbTable & "Thread(Thread_ID DESC);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX " & strDbTable & "Author" & strDbTable & "Thread ON " & strDbTable & "Thread(Author_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX " & strDbTable & "Topic" & strDbTable & "Thread ON " & strDbTable & "Thread(Topic_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE INDEX Topic_ID ON " & strDbTable & "Thread(Topic_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
					
			strSQL = "CREATE INDEX Guest_ID ON " & strDbTable & "GuestName(Guest_ID); "
			
			'Write to the database
			adoCon.Execute(strSQL)
			
	 		strSQL = "CREATE INDEX " & strDbTable & "Thread" & strDbTable & "GuestName ON " & strDbTable & "GuestName(Thread_ID);"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
	 		strSQL = "CREATE INDEX Thread_ID ON " & strDbTable & "GuestName(Thread_ID);"
			
			'Write to the database
			adoCon.Execute(strSQL)
			
			strSQL = "CREATE UNIQUE INDEX Session_ID ON " & strDbTable & "Session(Session_ID);"
			
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error Creating one or more Indexs <br />" & Replace(Err.description, "'", "\'") & ".';" & _
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
	
	
			'Stage 3 start
			Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
			vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br /><strong>Stage 3: Entering default values for new fields..... </strong>';" & _
			vbCrLf & "</script>")
			
			
			'Enter the default values in the LadderGroup Table
			'Primary Ladder Group
			strSQL = "INSERT INTO " & strDbTable & "LadderGroup ("
			strSQL = strSQL & "Ladder_Name "
			strSQL = strSQL & ") "
			strSQL = strSQL & "VALUES "
			strSQL = strSQL & "('Primary Ladder Group')"
			
			'Write to the database
			adoCon.Execute(strSQL)
			
			
			'Enter the default values in the UserGroup Table
			'Admin Group
			strSQL = "INSERT INTO " & strDbTable & "Group ("
			strSQL = strSQL & "Name, "
			strSQL = strSQL & "Minimum_posts, "
			strSQL = strSQL & "Special_rank, "
			strSQL = strSQL & "Ladder_ID, "
			strSQL = strSQL & "Stars, "
			strSQL = strSQL & "Starting_group "
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
			strSQL = "INSERT INTO " & strDbTable & "Group ("
			strSQL = strSQL & "Name, "
			strSQL = strSQL & "Minimum_posts, "
			strSQL = strSQL & "Special_rank, "
			strSQL = strSQL & "Ladder_ID, "
			strSQL = strSQL & "Stars, "
			strSQL = strSQL & "Starting_group "
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
			strSQL = "INSERT INTO " & strDbTable & "Group ("
			strSQL = strSQL & "Name, "
			strSQL = strSQL & "Minimum_posts, "
			strSQL = strSQL & "Special_rank, "
			strSQL = strSQL & "Ladder_ID, "
			strSQL = strSQL & "Stars, "
			strSQL = strSQL & "Starting_group "
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
			strSQL = "INSERT INTO " & strDbTable & "Group ("
			strSQL = strSQL & "Name, "
			strSQL = strSQL & "Minimum_posts, "
			strSQL = strSQL & "Special_rank, "
			strSQL = strSQL & "Ladder_ID, "
			strSQL = strSQL & "Stars, "
			strSQL = strSQL & "Starting_group "
			strSQL = strSQL & ") "
			strSQL = strSQL & "VALUES "
			strSQL = strSQL & "('Newbie', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'1', "
			strSQL = strSQL & "'1', "
			strSQL = strSQL & "'-1')"
			
			'Write to the database
			adoCon.Execute(strSQL)
			
			
			'Groupie Group
			strSQL = "INSERT INTO " & strDbTable & "Group ("
			strSQL = strSQL & "Name, "
			strSQL = strSQL & "Minimum_posts, "
			strSQL = strSQL & "Special_rank, "
			strSQL = strSQL & "Ladder_ID, "
			strSQL = strSQL & "Stars, "
			strSQL = strSQL & "Starting_group "
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
			strSQL = "INSERT INTO " & strDbTable & "Group ("
			strSQL = strSQL & "Name, "
			strSQL = strSQL & "Minimum_posts, "
			strSQL = strSQL & "Special_rank, "
			strSQL = strSQL & "Ladder_ID, "
			strSQL = strSQL & "Stars, "
			strSQL = strSQL & "Starting_group "
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
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error entering default values in the Table " & strDbTable & "Group<br />" & Replace(Err.description, "'", "\'") & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
			
			
			'Enter the default values in the Author Table
			'Enter the admin account into db
			strSQL = "INSERT INTO " & strDbTable & "Author ("
			strSQL = strSQL & "Group_ID, "
			strSQL = strSQL & "Username, "
			strSQL = strSQL & "User_code, "
			strSQL = strSQL & "Password, "
			strSQL = strSQL & "Salt, "
			strSQL = strSQL & "Show_email, "
			strSQL = strSQL & "Attach_signature, "
			strSQL = strSQL & "Time_offset, "
			strSQL = strSQL & "Time_offset_hours, "
			strSQL = strSQL & "Rich_editor, "
			strSQL = strSQL & "Date_format, "
			strSQL = strSQL & "Active, "
			strSQL = strSQL & "Reply_notify, "
			strSQL = strSQL & "PM_notify, "
			strSQL = strSQL & "No_of_posts, "
			strSQL = strSQL & "Signature, "
			strSQL = strSQL & "Join_date, "
			strSQL = strSQL & "Last_visit, "
			strSQL = strSQL & "Login_attempt, "
			strSQL = strSQL & "Banned, "
			strSQL = strSQL & "Info "
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
			strSQL = strSQL & "'-1', "
			strSQL = strSQL & "'dd/mm/yy', "
			strSQL = strSQL & "'-1', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'', "
			strSQL = strSQL & "'" & internationalDateTime(now()) & "', "
			strSQL = strSQL & "'" & internationalDateTime(now()) & "', "
			strSQL = strSQL & "'0',"
			strSQL = strSQL & "'0',"
			strSQL = strSQL & "'')"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'Enter the Guest account into db
			strSQL = "INSERT INTO " & strDbTable & "Author ("
			strSQL = strSQL & "Group_ID, "
			strSQL = strSQL & "Username, "
			strSQL = strSQL & "User_code, "
			strSQL = strSQL & "Password, "
			strSQL = strSQL & "Salt, "
			strSQL = strSQL & "Show_email, "
			strSQL = strSQL & "Attach_signature, "
			strSQL = strSQL & "Time_offset, "
			strSQL = strSQL & "Time_offset_hours, "
			strSQL = strSQL & "Rich_editor, "
			strSQL = strSQL & "Date_format, "
			strSQL = strSQL & "Active, "
			strSQL = strSQL & "Reply_notify, "
			strSQL = strSQL & "PM_notify, "
			strSQL = strSQL & "No_of_posts, "
			strSQL = strSQL & "Signature, "
			strSQL = strSQL & "Join_date, "
			strSQL = strSQL & "Last_visit, "
			strSQL = strSQL & "Login_attempt, "
			strSQL = strSQL & "Banned, "
			strSQL = strSQL & "Info "
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
			strSQL = strSQL & "'-1', "
			strSQL = strSQL & "'dd/mm/yy', "
			strSQL = strSQL & "'-1', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'0', "
			strSQL = strSQL & "'', "
			strSQL = strSQL & "'" & internationalDateTime(now()) & "', "
			strSQL = strSQL & "'" & internationalDateTime(now()) & "', "
			strSQL = strSQL & "'0',"
			strSQL = strSQL & "'0',"
			strSQL = strSQL & "'')"
	
			'Write to the database
			adoCon.Execute(strSQL)
	
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error entering default values in the Table " & strDbTable & "Author<br />" & Replace(Err.description, "'", "\'") & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
			
			
	
	
	
			'Enter the default values in the date time table
			strSQL = "INSERT INTO " & strDbTable & "DateTimeFormat ("
			strSQL = strSQL & "Date_format, "
			strSQL = strSQL & "Year_format, "
			strSQL = strSQL & "Seporator, "
			strSQL = strSQL & "Month1, "
			strSQL = strSQL & "Month2, "
			strSQL = strSQL & "Month3, "
			strSQL = strSQL & "Month4, "
			strSQL = strSQL & "Month5, "
			strSQL = strSQL & "Month6, "
			strSQL = strSQL & "Month7, "
			strSQL = strSQL & "Month8, "
			strSQL = strSQL & "Month9, "
			strSQL = strSQL & "Month10, "
			strSQL = strSQL & "Month11, "
			strSQL = strSQL & "Month12, "
			strSQL = strSQL & "Time_format, "
			strSQL = strSQL & "am, "
			strSQL = strSQL & "pm, "
			strSQL = strSQL & "Time_offset, "
			strSQL = strSQL & "Time_offset_hours "
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
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error entering default values in the Table " & strDbTable & "DateTimeFormat<br />" & Replace(Err.description, "'", "\'") & ".';" & _
				vbCrLf & "</script>")
	
				'Reset error object
				Err.Number = 0
	
				'Set the error boolean to True
				blnErrorOccured = True
			End If
	
	
	
			'Enter the default values in the smut table
			For intBadWordLoopCounter = 1 to 13
	
				'Write the SQL
				strSQL = "INSERT INTO " & strDbTable & "Smut (Smut, Word_replace) "
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
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />Error entering default values in the Table " & strDbTable & "Smut<br />" & Replace(Err.description, "'", "\'") & ".';" & _
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
			
			
			'If an error has occured write an error to the page
			If Err.Number <> 0 Then
				Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br/><br />Error creating default configuration<br />" & Replace(Err.description, "'", "\'") & ".';" & _
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
				vbCrLf & "document.getElementById('displayState').innerHTML = document.getElementById('displayState').innerHTML + '<br /><br />" & Replace(Err.description, "'", "\'") & "<br /><br /><h2>mySQL database is set up, but with Error!</h2>'" & _
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