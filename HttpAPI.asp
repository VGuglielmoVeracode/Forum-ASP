<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="functions/functions_format_post.asp" -->
<!--#include file="includes/emoticons_inc.asp" -->
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
Response.Buffer = true

'Make sure this page is not cached
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"



'If API is disabled
If blnHttpXmlApi = False Then 
	
	Call closeDatabase()
	
	'Reponse
	Response.Write("The Web Wiz Forums HTTP XML API is currently disabled.<br /><br />To enable the HTTP XML API login to the Forum Control Panel and enable it through the General Settings page.")
	
	'End the response	
	Response.Flush
	Response.End
End If





Dim strAPIversion	'Holds the version number of the API
Dim strApiAction	'Holds the API command
Dim strAdminUsername	'Holds the username of the admin account
Dim strAdminPassword	'Holds the passowrd of the admin account
Dim intErrorCode	'Holds the error code
Dim strErrorDescription	'Holds the error discription
Dim intRecordCount	'Holds the record count
Dim sarryRecords()	'Array holding all records returned for API method
Dim intRecordLoop
Dim strMemberName	'Holds the username 
Dim lngMemberID		'Holds the member ID
Dim lngTopicID
Dim intMaxResults	'Holds the max results to return
Dim strNewPassword	'New Password for member
Dim strSalt
Dim strMemberCode
Dim strUsername
Dim strPassword
Dim strEmail
Dim strRealName
Dim strGender
Dim strHomepage
Dim strSignature
Dim strICQNum
Dim blnShowEmail
Dim blnPMNotify
Dim blnAutoLogin
Dim blnUserActive 
Dim intUsersGroupID 
Dim lngPosts 
Dim strMemberTitle 
Dim blnSuspended 
Dim strAdminNotes 
Dim strAvatar
Dim strUserCode
Dim intForumStartingGroup
Dim blnNewsletter
Dim strEncryptedPassword
Dim lngPoints
Dim lngThanked




'API version
strAPIversion = "2.0"

'Intliase
intErrorCode = 0
lngMemberID = 0
intRecordLoop = 0
intMaxResults = 50

'Read in teh action to perform
strApiAction = Trim(Mid(Request("action"), 1, 25))


'If there is an action then run the page as XML
If strApiAction <> "" Then
	
	'Read in the admin username and password
	strAdminUsername = LCase(Trim(Mid(Request("Username"), 1, 20)))
	strAdminPassword = Trim(Mid(Request("Password"), 1, 20))


	'Set the response to XML
	Response.ContentType = "application/xml"
	
	'Set the top line of the page
	Response.Write("<?xml version=""1.0"" encoding=""" & strPageEncoding & """ ?>")
	
	
	
	'******  Admin Account Login Check *******

	'First checkout the username and password is OK
	'Get the master admin username and password from the db (Author_ID = 1), don't use the user imput in the SQL to prevent SQL injections
	strSQL = "SELECT " & strDbTable & "Author.Password, " & strDbTable & "Author.Salt, " & strDbTable & "Author.Username " & _
	"FROM " & strDbTable & "Author" & strDBNoLock & " " & _
	"WHERE " & strDbTable & "Author.Author_ID = 1;"
	
	'Set error trapping
	On Error Resume Next
		
	'Query the database
	rsCommon.Open strSQL, adoCon

	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then	Call errorMsg("An error has occurred while executing SQL query on database.", "get_master_admin_account", "HttpAPI.asp")
				
	'Disable error trapping
	On Error goto 0	
	
	
	'If no record returned then the main admin account has been removed
	If rsCommon.EOF Then
		
		'Reset Server Objects
		rsCommon.Close
		Call closeDatabase()
		
		Response.Write("" & _
		vbCrLf & "<ApiResponse>" & _
		vbCrLf & " <ErrorCode>-101</ErrorCode>" & _
		vbCrLf & " <ErrorDescription>Failed to retrieve Admin Login Data from Database</ErrorDescription>" & _
		vbCrLf & " <ResultData/>" & _
		vbCrLf & "</ApiResponse>")
		
		'End the response
		Response.Flush
		Response.End
	End If
	
	
	'Only encrypt password if this is enabled
	If blnEncryptedPasswords Then
			
		'Encrypt password so we can check it against the encypted password in the database
		'Read in the salt
		strAdminPassword = strAdminPassword & rsCommon("Salt")
	
		'Encrypt the entered password
		strAdminPassword = HashEncode(strAdminPassword)
	End If
	
	
	
	'If th admin username and password are incorrect return a fail
	If NOT strAdminUsername = LCase(rsCommon("Username")) OR NOT strAdminPassword = rsCommon("Password") Then
		
		
		'Reset Server Objects
		rsCommon.Close
		Call closeDatabase()
		
		Response.Write("" & _
		vbCrLf & "<ApiResponse>" & _
		vbCrLf & " <ErrorCode>-100</ErrorCode>" & _
		vbCrLf & " <ErrorDescription>Admin Login Fail</ErrorDescription>" & _
		vbCrLf & " <ResultData/>" & _
		vbCrLf & "</ApiResponse>")
		
		'End the response
		Response.Flush
		Response.End
		
	End If
	
	'Close recordset
	rsCommon.Close
	
	
	'Select API cation
	Select Case strApiAction
	
		'******  APIVersion *******
		Case "APIVersion" 
			
			ReDim Preserve sarryRecords(0)
			
			sarryRecords(0) =  vbCrLf & "   <ApiVersion>" & strApiVersion & "</ApiVersion>"
		
		
			
			
		'******  WebWizForumsVersion ******
		Case "WebWizForumsVersion"
			
			ReDim Preserve sarryRecords(0)
			
			sarryRecords(0) = ("" & _
			vbCrLf & "   <Software>Web Wiz Forums(TM)</Software>" & _
			vbCrLf & "   <Version>" & strVersion & "</Version>" & _
			vbCrLf & "   <ApiVersion>" & strApiVersion & "</ApiVersion>" & _
			vbCrLf & "   <Copyright>(C)2001-2019 Web Wiz Ltd. All rights reserved</Copyright>" & _
			vbCrLf & "   <BoardName>" & Server.HTMLEncode(strMainForumName) & "</BoardName>" & _
			vbCrLf & "   <URL>" & strForumPath & "</URL>" & _
			vbCrLf & "   <Email>" & strForumEmailAddress & "</Email>" & _
			vbCrLf & "   <Database>" & strDatabaseType & "</Database>" & _
			vbCrLf & "   <InstallID>" & strInstallID & "</InstallID>" & _
			vbCrLf & "   <NewsPad>" & blnWebWizNewsPad & "</NewsPad>" & _
			vbCrLf & "   <NewsPadURL>" & strWebWizNewsPadURL & "</NewsPadURL>")
			
		
		
		
		
		
		
		'******  GetMemberByName OR GetMemberByID ******
		Case "GetMemberByName", "GetMemberByID"
			
			If strApiAction = "GetMemberByName" Then
				'Read in username
				strMemberName = Trim(Mid(Request("MemberName"), 1, 20))
				strMemberName = formatSQLInput(strMemberName)
			Else
				If isNumeric(Request("MemberID")) Then
				
					lngMemberID =  LngC(Request("MemberID"))
				Else
					lngMemberID = -1
				End If
			End If
			
			'SQL
			strSQL = "SELECT " & strDbTable & "Author.*, " & strDbTable & "Group.*, " & strDbTable & "Group.Group_ID AS GroupID " & _
			"FROM " & strDbTable & "Author" & strDBNoLock & ", " & strDbTable & "Group" & strDBNoLock & " " & _
			"WHERE " & strDbTable & "Author.Group_ID = " & strDbTable & "Group.Group_ID "
			If strApiAction = "GetMemberByName" Then
				strSQL = strSQL & "AND " & strDbTable & "Author.Username = '" & strMemberName & "'; "
			Else
				strSQL = strSQL & "AND " & strDbTable & "Author.Author_ID = " & lngMemberID & ";"
			End If
			
			'Query the database
			rsCommon.Open strSQL, adoCon
			
			'If nothing returned then an error
			If rsCommon.EOF Then
				
				intErrorCode = -150
				strErrorDescription = "Member not found"
			
			'Else member is found so write XML	
			Else
				
				
				ReDim Preserve sarryRecords(0)
				
				'Read in real name here incase it contains a null
				If isNull(rsCommon("Real_name")) Then strRealName = "" Else strRealName = rsCommon("Real_name")
				
				sarryRecords(0) = ("" & _
				vbCrLf & "   <Username>" & Server.HTMLEncode(rsCommon("Username")) & "</Username>" & _
				vbCrLf & "   <UserID>" & rsCommon("Author_ID") & "</UserID>" & _
				vbCrLf & "   <Group>" & Server.HTMLEncode(rsCommon("Name")) & "</Group>" & _
				vbCrLf & "   <GroupID>" & rsCommon("GroupID") & "</GroupID>" & _
				vbCrLf & "   <MemberCode>" & rsCommon("User_code") & "</MemberCode>")
				If blnEncryptedPasswords Then	
					sarryRecords(0) = sarryRecords(0) & ("" & _
					vbCrLf & "   <EncryptedPassword>" & rsCommon("Password") & "</EncryptedPassword>" & _
					vbCrLf & "   <Salt>" & rsCommon("Salt") & "</Salt>")
				Else
					sarryRecords(0) = sarryRecords(0) & ("" & _
					vbCrLf & "   <Password>" & rsCommon("Password") & "</Password>")
				End If	
				sarryRecords(0) = sarryRecords(0) & ("" & _
				vbCrLf & "   <Active>" & CBool(rsCommon("Active")) & "</Active>" & _
				vbCrLf & "   <Suspended>" & CBool(rsCommon("Banned"))  & "</Suspended>")
				If isDate(rsCommon("Join_date")) Then sarryRecords(0) = sarryRecords(0) & vbCrLf & "   <Joined>" & internationalDateTime(CDate(rsCommon("Join_date"))) & "</Joined>" Else sarryRecords(0) = sarryRecords(0) & vbCrLf & "   <Joined/>"
				If isDate(rsCommon("Last_visit")) Then sarryRecords(0) = sarryRecords(0) & vbCrLf & "   <LastVisit>" & internationalDateTime(CDate(rsCommon("Last_visit"))) & "</LastVisit>" Else sarryRecords(0) = sarryRecords(0) & vbCrLf & "   <LastVisit/>"
				sarryRecords(0) = sarryRecords(0) & ("" & _
				vbCrLf & "   <Email>" & rsCommon("Author_email") & "</Email>" & _
				vbCrLf & "   <Name>" & Server.HTMLEncode(strRealName) & "</Name>")
				If isDate(rsCommon("DOB")) Then sarryRecords(0) = sarryRecords(0) & vbCrLf & "   <DOB>" & internationalDateTime(CDate(rsCommon("DOB"))) & "</DOB>" Else sarryRecords(0) = sarryRecords(0) & vbCrLf & "   <DOB/>"
				If isNull(rsCommon("Gender")) = False Then sarryRecords(0) = sarryRecords(0) & vbCrLf & "   <Gender>" & Server.HTMLEncode(rsCommon("Gender")) & "</Gender>" Else sarryRecords(0) = sarryRecords(0) & vbCrLf & "   <Gender/>"
				sarryRecords(0) = sarryRecords(0) & ("" & _
				vbCrLf & "   <PostCount>" & rsCommon("No_of_posts") & "</PostCount>" & _
				vbCrLf & "   <Points>" & rsCommon("Points") & "</Points>" & _
				vbCrLf & "   <Thanked>" & rsCommon("Thanked") & "</Thanked>" & _
				vbCrLf & "   <Newsletter>" & CBool(rsCommon("Newsletter")) & "</Newsletter>")
			
			End If
			
			'Reset Server Objects
			rsCommon.Close
			
			
			
			
			
		
		'******  ActivateMember  ******
		Case "ActivateMember"
			
			
			'Read in username
			strMemberName = Trim(Mid(Request("MemberName"), 1, 20))
			strMemberName = formatSQLInput(strMemberName)
			
			
			'SQL
			strSQL = "SELECT " & strDbTable & "Author.Author_ID, " & strDbTable & "Author.Username, " & strDbTable & "Author.Active " & _
			"FROM " & strDbTable & "Author" & strDBNoLock & " " & _
			"WHERE " & strDbTable & "Author.Username = '" & strMemberName & "'; "
			
			'Query the database
			rsCommon.Open strSQL, adoCon
			
			'If nothing returned then an error
			If rsCommon.EOF Then
				
				intErrorCode = -150
				strErrorDescription = "Member not found"
			
			'Else member is found so write XML	
			Else
				ReDim Preserve sarryRecords(0)
				
				'Update user status to active
				strSQL = "UPDATE " & strDbTable & "Author" & strRowLock & " " & _
				"SET " & strDbTable & "Author.Active = " & strDBTrue & " " & _
				"WHERE " & strDbTable & "Author.Username = '" & strMemberName & "'; "
				
				'Write to the database
				adoCon.Execute(strSQL)
				
				
				sarryRecords(0) = ("" & _
				vbCrLf & "   <Username>" & Server.HTMLEncode(rsCommon("Username")) & "</Username>" & _
				vbCrLf & "   <UserID>" & rsCommon("Author_ID") & "</UserID>" & _
				vbCrLf & "   <Active>True</Active>")
			
			End If
			
			'Reset Server Objects
			rsCommon.Close
		
		
		
		
		'******  SuspendMember  OR UnsubspendMember ******
		Case "SuspendMember", "UnsubspendMember"
			
			
			'Read in username
			strMemberName = Trim(Mid(Request("MemberName"), 1, 20))
			strMemberName = formatSQLInput(strMemberName)
			
			
			'SQL
			strSQL = "SELECT " & strDbTable & "Author.Author_ID, " & strDbTable & "Author.Username, " & strDbTable & "Author.Banned " & _
			"FROM " & strDbTable & "Author" & strDBNoLock & " " & _
			"WHERE " & strDbTable & "Author.Username = '" & strMemberName & "'; "
			
			'Query the database
			rsCommon.Open strSQL, adoCon
			
			'If nothing returned then an error
			If rsCommon.EOF Then
				
				intErrorCode = -150
				strErrorDescription = "Member not found"
			
			'Else member is found so write XML	
			Else
				ReDim Preserve sarryRecords(0)
				
				
				If strApiAction = "UnsubspendMember" Then
					
					'Update db
					strSQL = "UPDATE " & strDbTable & "Author" & strRowLock & " " & _
					"SET " & strDbTable & "Author.Banned = " & strDBFalse & " " & _
					"WHERE " & strDbTable & "Author.Username = '" & strMemberName & "'; "
					
					'Write to the database
					adoCon.Execute(strSQL)
					
					
					sarryRecords(0) = ("" & _
					vbCrLf & "   <Username>" & Server.HTMLEncode(rsCommon("Username")) & "</Username>" & _
					vbCrLf & "   <UserID>" & rsCommon("Author_ID") & "</UserID>" & _
					vbCrLf & "   <Suspended>False</Suspended>")
					
				Else
					'Update db
					strSQL = "UPDATE " & strDbTable & "Author" & strRowLock & " " & _
					"SET " & strDbTable & "Author.Banned = " & strDBTrue & " " & _
					"WHERE " & strDbTable & "Author.Username = '" & strMemberName & "'; "
					
					'Write to the database
					adoCon.Execute(strSQL)
					
					
					sarryRecords(0) = ("" & _
					vbCrLf & "   <Username>" & Server.HTMLEncode(rsCommon("Username")) & "</Username>" & _
					vbCrLf & "   <UserID>" & rsCommon("Author_ID") & "</UserID>" & _
					vbCrLf & "   <Suspended>True</Suspended>")
				End If
			
			End If
			
			'Reset Server Objects
			rsCommon.Close
		
		
		
		
		
		'******  GetForums  ******
		Case "GetForums"
			
			
			'SQL
			strSQL = "SELECT " & strDbTable & "Category.*, " & strDbTable & "Forum.* " & _
			"FROM " & strDbTable & "Category" & strDBNoLock & ", " & strDbTable & "Forum" & strDBNoLock & " " & _
			"WHERE " & strDbTable & "Category.Cat_ID = " & strDbTable & "Forum.Cat_ID " & _
			"ORDER BY " & strDbTable & "Category.Cat_order, " & strDbTable & "Forum.Forum_Order;"
			
			'Query the database
			rsCommon.Open strSQL, adoCon
			
			'If nothing returned then an error
			If rsCommon.EOF Then
				
				intErrorCode = -150
				strErrorDescription = "No Forums Found"
			
			'Else forums are found so write XML	
			Else
				
				'Loop through records
				DO WHILE NOT rsCommon.EOF
					
					ReDim Preserve sarryRecords(intRecordLoop)
					
					
					sarryRecords(intRecordLoop) = ("" & _
					vbCrLf & "   <ForumName>" & Server.HTMLEncode(rsCommon("Forum_name")) & "</ForumName>" & _
					vbCrLf & "   <ForumID>" & rsCommon("Forum_ID") & "</ForumID>" & _
					vbCrLf & "   <SubForumID>" & rsCommon("Sub_ID") & "</SubForumID>" & _
					vbCrLf & "   <CatName>" & rsCommon("Cat_Name") & "</CatName>" & _
					vbCrLf & "   <CatID>" & rsCommon("Cat_ID") & "</CatID>" & _
					vbCrLf & "   <ForumDescription><![CDATA[" & rsCommon("Forum_description") & "]]></ForumDescription>")
					If rsCommon("Password") <> "" Then sarryRecords(intRecordLoop) = sarryRecords(intRecordLoop) & vbCrLf & "   <Password>True</Password>" Else sarryRecords(intRecordLoop) = sarryRecords(intRecordLoop) & vbCrLf & "   <Password>False</Password>"
					sarryRecords(intRecordLoop) = sarryRecords(intRecordLoop) & ("" & _
					vbCrLf & "   <Locked>" & CBool(rsCommon("Locked")) & "</Locked>" & _
					vbCrLf & "   <Hidden>" & CBool(rsCommon("Hide")) & "</Hidden>" & _
					vbCrLf & "   <TopicCount>" & rsCommon("No_of_topics") & "</TopicCount>" & _
					vbCrLf & "   <PostCount>" & rsCommon("No_of_posts") & "</PostCount>" & _
					vbCrLf & "   <LastPostMemberID>" & rsCommon("Last_post_author_ID") & "</LastPostMemberID>")
					If isDate(rsCommon("Last_post_date")) Then sarryRecords(intRecordLoop) = sarryRecords(intRecordLoop) & vbCrLf & "   <LastPostDate>" & internationalDateTime(CDate(rsCommon("Last_post_date"))) & "</LastPostDate>" Else sarryRecords(intRecordLoop) = sarryRecords(intRecordLoop) & vbCrLf & "   <LastPostDate/>"
					sarryRecords(intRecordLoop) = sarryRecords(intRecordLoop) & ("" & _
					vbCrLf & "   <LastPostTopicID>" & rsCommon("Last_topic_ID") & "</LastPostTopicID>")
					
					intRecordLoop = intRecordLoop + 1
					
					'Move to next record
					rsCommon.MoveNext
				Loop
			
			End If
			
			'Reset Server Objects
			rsCommon.Close
			
			
		
		
		'******  LockForumByID   OR  UnLockForumByID ******
		Case "LockForumByID", "UnLockForumByID"
			
			
			'Read in forum ID
			intForumID = IntC(Request("ForumID"))
			
			
			'SQL
			strSQL = "SELECT " & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Forum_name, " & strDbTable & "Forum.Locked " & _
			"FROM " & strDbTable & "Forum" & strDBNoLock & " " & _
			"WHERE " & strDbTable & "Forum.Forum_ID = " & intForumID & "; "
			
			'Query the database
			rsCommon.Open strSQL, adoCon
			
			'If nothing returned then an error
			If rsCommon.EOF Then
				
				intErrorCode = -150
				strErrorDescription = "Forum not found"
			
			'Else forum is found so write XML	
			Else
				ReDim Preserve sarryRecords(0)
				
				'Update db
				If strApiAction = "UnLockForumByID"  Then 
					
					'Update user status to active
					strSQL = "UPDATE " & strDbTable & "Forum" & strRowLock & " " & _
					"SET " & strDbTable & "Forum.Locked = " & strDBFalse & " " & _
					"WHERE " & strDbTable & "Forum.Forum_ID = " & intForumID & "; "
					
					'Write to the database
					adoCon.Execute(strSQL)
					
					sarryRecords(0) = ("" & _
					vbCrLf & "   <ForumID>" & Server.HTMLEncode(rsCommon("Forum_ID")) & "</ForumID>" & _
					vbCrLf & "   <ForumDescription>" & Server.HTMLEncode(rsCommon("Forum_name")) & "</ForumDescription>" & _
					vbCrLf & "   <Locked>False</Locked>")
				
				Else
					strSQL = "UPDATE " & strDbTable & "Forum" & strRowLock & " " & _
					"SET " & strDbTable & "Forum.Locked = " & strDBTrue & " " & _
					"WHERE " & strDbTable & "Forum.Forum_ID = " & intForumID & "; "
					
					'Write to the database
					adoCon.Execute(strSQL)
					
					sarryRecords(0) = ("" & _
					vbCrLf & "   <ForumID>" & Server.HTMLEncode(rsCommon("Forum_ID")) & "</ForumID>" & _
					vbCrLf & "   <ForumDescription>" & Server.HTMLEncode(rsCommon("Forum_name")) & "</ForumDescription>" & _
					vbCrLf & "   <Locked>True</Locked>")
				
				End If
			
			End If
			
			'Reset Server Objects
			rsCommon.Close
		
		
		
		'******  GetTopicNameByID  OR CloseTopicByID OR OpenTopicByID ******
		Case "GetTopicNameByID", "CloseTopicByID", "OpenTopicByID"
			
			'Read in the TopicID
			lngTopicID = LngC(Request("TopicID"))
			
			'SQL
			strSQL = "SELECT " & strDbTable & "Topic.* " & _
			"FROM " & strDbTable & "Topic" & strDBNoLock & " " & _
			"WHERE " & strDbTable & "Topic.Topic_ID = " & lngTopicID & ";"
			
			'Query the database
			rsCommon.Open strSQL, adoCon
			
			'If nothing returned then an error
			If rsCommon.EOF Then
				
				intErrorCode = -150
				strErrorDescription = "Topic not found"
			
			'Else forums are found so write XML	
			Else
				
					
				ReDim Preserve sarryRecords(0)
				
				
				
				'Update db
				If strApiAction = "CloseTopicByID"  Then 
					
					'Update db
					strSQL = "UPDATE " & strDbTable & "Topic" & strRowLock & " " & _
					"SET " & strDbTable & "Topic.Locked = " & strDBTrue & " " & _
					"WHERE " & strDbTable & "Topic.Topic_ID = " & lngTopicID & "; "
					
					'Write to the database
					adoCon.Execute(strSQL)
					
					'Rerun recordset query
					rsCommon.ReQuery
				
				ElseIf strApiAction = "OpenTopicByID"  Then
					
					'Update db
					strSQL = "UPDATE " & strDbTable & "Topic" & strRowLock & " " & _
					"SET " & strDbTable & "Topic.Locked = " & strDBFalse & " " & _
					"WHERE " & strDbTable & "Topic.Topic_ID = " & lngTopicID & "; "
					
					'Write to the database
					adoCon.Execute(strSQL)
					
					'Rerun recordset query
					rsCommon.ReQuery
				End If
				
				
				
					
					
				sarryRecords(intRecordLoop) = ("" & _
				vbCrLf & "   <TopicName><![CDATA[" & rsCommon("Subject") & "]]></TopicName>" & _
				vbCrLf & "   <TopicID>" & rsCommon("Topic_ID") & "</TopicID>" & _
				vbCrLf & "   <ForumID>" & rsCommon("Forum_ID") & "</ForumID>" & _
				vbCrLf & "   <TopicLocked>" & CBool(rsCommon("Locked")) & "</TopicLocked>" & _
				vbCrLf & "   <Hidden>" & CBool(rsCommon("Hide")) & "</Hidden>" & _
				vbCrLf & "   <ReplyCount>" & rsCommon("No_of_replies") & "</ReplyCount>")
				If isDate(rsCommon("Event_date")) Then sarryRecords(0) = sarryRecords(0) & vbCrLf & "   <EventDateStart>" & internationalDateTime(CDate(rsCommon("Event_date"))) & "</EventDateStart>" Else sarryRecords(0) = sarryRecords(0) & vbCrLf & "   <EventDateStart/>"
				If isDate(rsCommon("Event_date_end")) Then sarryRecords(0) = sarryRecords(0) & vbCrLf & "   <EventDateEnd>" & internationalDateTime(CDate(rsCommon("Event_date_end"))) & "</EventDateEnd>" Else sarryRecords(0) = sarryRecords(0) & vbCrLf & "   <EventDateEnd/>"
				sarryRecords(intRecordLoop) = sarryRecords(intRecordLoop) & ("" & _
				vbCrLf & "   <ViewCount>" & rsCommon("No_of_views") & "</ViewCount>")
			
			End If
			
			'Reset Server Objects
			rsCommon.Close
			
			
		
		
		
		'******  GetLastTopics  OR GetLastTopicsByForumID  ******
		Case "GetLastTopics", "GetLastTopicsByForumID" 
			
			'Read in the max results
			If isNumeric(Request("MaxResults")) Then
				
				'Get the max results to show, trim this to a 3 figure number, as only 50 allowed, and prevent errors
				intMaxResults = Trim(Mid(Request("MaxResults"), 1, 3))
				
				'Convert into integer
				intMaxResults = IntC(intMaxResults)
				
				'Set some defaults if out of range
				If intMaxResults > 50 Then intMaxResults = 50
				If intMaxResults < 1 Then intMaxResults = 1
			End If
			
			'If GetLastPostsByForumID then read in the forum ID
			If strApiAction = "GetLastTopicsByForumID" Then
				If isNumeric(Request("ForumID")) Then
					intForumID =  LngC(Request("ForumID"))
				Else
					intForumID = -1
				End If
				
			End If
			
			
			'SQL
			strSQL = "" & _
			"SELECT "
			If strDatabaseType = "SQLServer" OR strDatabaseType = "Access" Then
				strSQL = strSQL & " TOP " & intMaxResults & " "
			End If
			strSQL = strSQL & _
			"" & strDbTable & "Forum.Forum_name, " & strDbTable & "Topic.* " & _
			"FROM " & strDbTable & "Forum, " & strDbTable & "Topic " & _
			"WHERE " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Topic.Forum_ID "
				
			'If looking at a forum only, only get posts from that forum
			If intForumID <> 0 Then strSQL = strSQL & "AND " & strDbTable & "Topic.Forum_ID = " & intForumID & " "
			
			
			strSQL = strSQL & "AND (" & strDbTable & "Topic.Hide = " & strDBFalse & ") " & _
			"ORDER BY " & strDbTable & "Topic.Last_Thread_ID DESC"
			
			'mySQL limit operator
			If strDatabaseType = "mySQL" Then
				strSQL = strSQL & " LIMIT " & intMaxResults
			End If
			
			'Query the database
			rsCommon.Open strSQL, adoCon
			
			'If nothing returned then an error
			If rsCommon.EOF Then
				
				intErrorCode = -150
				If strApiAction = "GetLastTopicsByForumID" Then
					strErrorDescription = "Forum not found or no topics in forum"
				Else
					strErrorDescription = "No topics found"
				End If
			
			'Else forums are found so write XML	
			Else
				
				'Loop through records
				DO WHILE NOT rsCommon.EOF
					
					ReDim Preserve sarryRecords(intRecordLoop)
					
					sarryRecords(intRecordLoop) = ("" & _
					vbCrLf & "   <TopicName><![CDATA[" & rsCommon("Subject") & "]]></TopicName>" & _
					vbCrLf & "   <TopicID>" & rsCommon("Topic_ID") & "</TopicID>" & _
					vbCrLf & "   <ForumName>" & Server.HTMLEncode(rsCommon("Forum_name")) & "</ForumName>" & _
					vbCrLf & "   <ForumID>" & rsCommon("Forum_ID") & "</ForumID>" & _
					vbCrLf & "   <TopicLocked>" & CBool(rsCommon("Locked")) & "</TopicLocked>" & _
					vbCrLf & "   <Priority>" & rsCommon("Priority") & "</Priority>" & _
					vbCrLf & "   <Hidden>" & CBool(rsCommon("Hide")) & "</Hidden>" & _
					vbCrLf & "   <ReplyCount>" & rsCommon("No_of_replies") & "</ReplyCount>")
					If isDate(rsCommon("Event_date")) Then sarryRecords(intRecordLoop) = sarryRecords(intRecordLoop) & vbCrLf & "   <EventDateStart>" & internationalDateTime(CDate(rsCommon("Event_date"))) & "</EventDateStart>" Else sarryRecords(intRecordLoop) = sarryRecords(intRecordLoop) & vbCrLf & "   <EventDateStart/>"
					If isDate(rsCommon("Event_date_end")) Then sarryRecords(intRecordLoop) = sarryRecords(intRecordLoop) & vbCrLf & "   <EventDateEnd>" & internationalDateTime(CDate(rsCommon("Event_date_end"))) & "</EventDateEnd>" Else sarryRecords(intRecordLoop) = sarryRecords(intRecordLoop) & vbCrLf & "   <EventDateEnd/>"
					sarryRecords(intRecordLoop) = sarryRecords(intRecordLoop) & ("" & _
					vbCrLf & "   <ViewCount>" & rsCommon("No_of_views") & "</ViewCount>" & _
					vbCrLf & "   <Rating>" & rsCommon("Rating") & "</Rating>")
					
					intRecordLoop = intRecordLoop + 1
					
					'Move to next record
					rsCommon.MoveNext
				Loop
			
			End If
			
			'Reset Server Objects
			rsCommon.Close
		
		
				
		
			
			
			
		'******  GetLastPosts  OR GetLastPostsByForumID  ******
		Case "GetLastPosts", "GetLastPostsByForumID"
			
			'Read in the max results
			If isNumeric(Request("MaxResults")) Then
				
				'Get the max results to show, trim this to a 3 figure number, as only 50 allowed, and prevent errors
				intMaxResults = Trim(Mid(Request("MaxResults"), 1, 3))
				
				'Convert into integer
				intMaxResults = IntC(intMaxResults)
				
				'Set some defaults if out of range
				If intMaxResults > 50 Then intMaxResults = 50
				If intMaxResults < 1 Then intMaxResults = 1
			End If
			
			'If GetLastPostsByForumID then read in the forum ID
			If strApiAction = "GetLastPostsByForumID" Then
				If isNumeric(Request("ForumID")) Then
					intForumID =  LngC(Request("ForumID"))
				Else
					intForumID = -1
				End If
				
			End If
			
			
			'SQL
			strSQL = "" & _
			"SELECT "
			If strDatabaseType = "SQLServer" OR strDatabaseType = "Access" Then
				strSQL = strSQL & " TOP " & intMaxResults & " "
			End If
			strSQL = strSQL & _
			"" & strDbTable & "Forum.Forum_name, " & strDbTable & "Topic.*, " & strDbTable & "Thread.Thread_ID, " & strDbTable & "Thread.Message_date, " & strDbTable & "Thread.Hide, " & strDbTable & "Thread.Thanks, " & strDbTable & "Thread.Answer, " & strDbTable & "Thread.IP_addr, " & strDbTable & "Author.Author_ID, " & strDbTable & "Author.Username, " & strDbTable & "Thread.Message  " & _
			"FROM " & strDbTable & "Forum, " & strDbTable & "Topic, " & strDbTable & "Author, " & strDbTable & "Thread " & _
			"WHERE " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Topic.Forum_ID " & _
				"AND " & strDbTable & "Topic.Topic_ID = " & strDbTable & "Thread.Topic_ID " & _
				"AND " & strDbTable & "Author.Author_ID = " & strDbTable & "Thread.Author_ID "
			
			'If looking at a forum only, only get posts from tha forum
			If intForumID <> 0 Then strSQL = strSQL & "AND " & strDbTable & "Topic.Forum_ID = " & intForumID & " "
			
			
			strSQL = strSQL & "AND (" & strDbTable & "Topic.Hide = " & strDBFalse & " AND " & strDbTable & "Thread.Hide = " & strDBFalse & ") " & _
			"ORDER BY " & strDbTable & "Thread.Thread_ID DESC"
			
			'mySQL limit operator
			If strDatabaseType = "mySQL" Then
				strSQL = strSQL & " LIMIT " & intMaxResults
			End If
			
			'Query the database
			rsCommon.Open strSQL, adoCon
			
			'If nothing returned then an error
			If rsCommon.EOF Then
				
				intErrorCode = -150
				If strApiAction = "GetLastPostsByForumID" Then
					strErrorDescription = "Forum not found or no posts in forum"
				Else
					strErrorDescription = "No posts found"
				End If
			
			'Else forums are found so write XML	
			Else
				
				'Loop through records
				DO WHILE NOT rsCommon.EOF
					
					ReDim Preserve sarryRecords(intRecordLoop)
					
					sarryRecords(intRecordLoop) = ("" & _
					vbCrLf & "   <TopicName><![CDATA[" & rsCommon("Subject") & "]]></TopicName>" & _
					vbCrLf & "   <TopicID>" & rsCommon("Topic_ID") & "</TopicID>" & _
					vbCrLf & "   <ForumName>" & Server.HTMLEncode(rsCommon("Forum_name")) & "</ForumName>" & _
					vbCrLf & "   <ForumID>" & rsCommon("Forum_ID") & "</ForumID>" & _
					vbCrLf & "   <TopicLocked>" & CBool(rsCommon("Locked")) & "</TopicLocked>" & _
					vbCrLf & "   <ReplyCount>" & rsCommon("No_of_replies") & "</ReplyCount>")
					If isDate(rsCommon("Event_date")) Then sarryRecords(intRecordLoop) = sarryRecords(intRecordLoop) & vbCrLf & "   <EventDateStart>" & internationalDateTime(CDate(rsCommon("Event_date"))) & "</EventDateStart>" Else sarryRecords(intRecordLoop) = sarryRecords(intRecordLoop) & vbCrLf & "   <EventDateStart/>"
					If isDate(rsCommon("Event_date_end")) Then sarryRecords(intRecordLoop) = sarryRecords(intRecordLoop) & vbCrLf & "   <EventDateEnd>" & internationalDateTime(CDate(rsCommon("Event_date_end"))) & "</EventDateEnd>" Else sarryRecords(intRecordLoop) = sarryRecords(intRecordLoop) & vbCrLf & "   <EventDateEnd/>"
					sarryRecords(intRecordLoop) = sarryRecords(intRecordLoop) & ("" & _
					vbCrLf & "   <Thanks>" & rsCommon("Thanks") & "</Thanks>" & _
					vbCrLf & "   <Answer>" & CBool(rsCommon("Answer")) & "</Answer>" & _
					vbCrLf & "   <Hidden>" & CBool(rsCommon("" & strDbTable & "Thread.Hide")) & "</Hidden>" & _
					vbCrLf & "   <PostID>" & rsCommon("Thread_ID") & "</PostID>")
					If isDate(rsCommon("Message_date")) Then sarryRecords(intRecordLoop) = sarryRecords(intRecordLoop) & vbCrLf & "   <PostDate>" & internationalDateTime(CDate(rsCommon("Message_date"))) & "</PostDate>" Else sarryRecords(intRecordLoop) = sarryRecords(intRecordLoop) & vbCrLf & "   <PostDate/>"
					sarryRecords(intRecordLoop) = sarryRecords(intRecordLoop) & ("" & _
					vbCrLf & "   <MemberID>" & rsCommon("Author_ID") & "</MemberID>" & _
					vbCrLf & "   <MemberName>" & Server.HTMLEncode(rsCommon("Username")) & "</MemberName>" & _
					vbCrLf & "   <MemberIP>" & rsCommon("IP_addr") & "</MemberIP>" & _
					vbCrLf & "   <Post><![CDATA[" & rsCommon("Message") & "]]></Post>")
					
					intRecordLoop = intRecordLoop + 1
					
					'Move to next record
					rsCommon.MoveNext
				Loop
			
			End If
			
			'Reset Server Objects
			rsCommon.Close
			
		
		
		
		'******  ChangeMemberPassword  ******
		Case "ChangeMemberPassword"
			
			
			'Read in username
			strMemberName = Trim(Mid(Request("MemberName"), 1, 20))
			strMemberName = formatSQLInput(strMemberName)
			
			'Read in password
			strNewPassword = LCase(Trim(Mid(Request("NewPassword"), 1, 15)))
			
			
			'SQL
			strSQL = "SELECT " & strDbTable & "Author.Author_ID, " & strDbTable & "Author.Username, " & strDbTable & "Author.Password, " & strDbTable & "Author.Salt " & _
			"FROM " & strDbTable & "Author" & strDBNoLock & " " & _
			"WHERE " & strDbTable & "Author.Username = '" & strMemberName & "'; "
			
			'Query the database
			rsCommon.Open strSQL, adoCon
			
			'If nothing returned then an error
			If rsCommon.EOF Then
				
				intErrorCode = -150
				strErrorDescription = "Member not found"
			
			'If no password 
			ElseIf Len(strNewPassword) < 2 Then
				
				intErrorCode = -200
				strErrorDescription = "Password length is to short"
			
			'Else member is found so write XML	
			Else
				ReDim Preserve sarryRecords(0)
				
				 'Encrypt password
				If strNewPassword <> "" Then
					
					'Encrypt password
					If blnEncryptedPasswords Then																							
				
						'Genrate a slat value
					       	strSalt = getSalt(Len(strNewPassword))
					
					       'Concatenate salt value to the password
					       strNewPassword = strNewPassword & strSalt
					
					       'Encrypt the password
					       strNewPassword = HashEncode(strNewPassword)
					
					'Else the password is not set to be encrypted so make sure it is SQL safe
					Else
				
						strNewPassword = formatSQLInput(strNewPassword)
					End If
				End If
				
				'Generate new usercode for user
				strMemberCode = userCode(strMemberName)
				
					
				'Update db
				strSQL = "UPDATE " & strDbTable & "Author" & strRowLock & " " & _
				"SET " & _
				strDbTable & "Author.User_code = '" & strMemberCode & "', " & _
				strDbTable & "Author.Password = '" & strNewPassword & "', " & _
				strDbTable & "Author.Salt = '" & strSalt & "' " & _
				"WHERE " & strDbTable & "Author.Username = '" & strMemberName & "'; "
					
				'Write to the database
				adoCon.Execute(strSQL)
					
					
				sarryRecords(0) = ("" & _
				vbCrLf & "   <Username>" & Server.HTMLEncode(rsCommon("Username")) & "</Username>" & _
				vbCrLf & "   <UserID>" & rsCommon("Author_ID") & "</UserID>")
				If blnEncryptedPasswords Then	
					sarryRecords(0) = sarryRecords(0) & ("" & _
					vbCrLf & "   <EncryptedPassword>" & strNewPassword & "</EncryptedPassword>" & _
					vbCrLf & "   <Salt>" & strSalt & "</Salt>")
				Else
					sarryRecords(0) = sarryRecords(0) & ("" & _
					vbCrLf & "   <Password>" & strNewPassword & "</Password>")
				End If	
				sarryRecords(0) = sarryRecords(0) & ("" & _
				vbCrLf & "   <MemberCode>" & strMemberCode & "</MemberCode>")
				
			
			End If
			
			'Reset Server Objects
			rsCommon.Close
			
		
		
		
		
		
		
		'******  LoginMemberCookie  ******
		Case "LoginMemberCookie"
			
			
			'Read in username
			strMemberName = Trim(Mid(Request("MemberName"), 1, 20))
			strMemberName = formatSQLInput(strMemberName)
			
			'Read in password
			strPassword = Trim(Mid(Request("MemberPassword"), 1, 15))
			
			
			'SQL
			strSQL = "SELECT " & strDbTable & "Author.Author_ID, " & strDbTable & "Author.Username, " & strDbTable & "Author.User_code, " & strDbTable & "Author.Password, " & strDbTable & "Author.Salt " & _
			"FROM " & strDbTable & "Author" & strDBNoLock & " " & _
			"WHERE " & strDbTable & "Author.Username = '" & strMemberName & "'; "
			
			'Query the database
			rsCommon.Open strSQL, adoCon
			
			'If nothing returned then an error
			If rsCommon.EOF Then
				
				intErrorCode = -150
				strErrorDescription = "Member not found"
			
			'Else member is found so write XML	
			Else
				
				'If password is enetred then check it
				If strPassword <> "" Then
					
					'Only encrypt password if this is enabled
					If blnEncryptedPasswords Then
						
						'Encrypt password so we can check it against the encypted password in the database
						'Read in the salt
						strPassword = strPassword & rsCommon("Salt")
				
						'Encrypt the entered password
						strPassword = HashEncode(strPassword)
					End If
					
					'If password is wrong then tell the user
					If strPassword <> rsCommon("Password") Then
						
						intErrorCode = -160
						strErrorDescription = "Member password incorrect"
					End If
				End If
			
				'If no error from the password check then display login details for the user
				If intErrorCode = 0 Then
				
					ReDim Preserve sarryRecords(0)
						
					sarryRecords(0) = ("" & _
					vbCrLf & "   <Username>" & Server.HTMLEncode(rsCommon("Username")) & "</Username>" & _
					vbCrLf & "   <UserID>" & rsCommon("Author_ID") & "</UserID>" & _
					vbCrLf & "   <CookieName>" & strCookiePrefix & "sLID</CookieName>" & _
					vbCrLf & "   <CookieKey>UID</CookieKey>" & _
					vbCrLf & "   <CookieData>" &  Server.HTMLEncode(rsCommon("User_code")) & "</CookieData>" & _
					vbCrLf & "   <CookiePath>" & strCookiePath & "</CookiePath>" & _
					vbCrLf & "   <CookieDomain>" & strCookieDomain & "</CookieDomain>" & _
					vbCrLf & "   <ForumPath>" & strForumPath & "</ForumPath>")
				End If
			
			End If
			
			'Reset Server Objects
			rsCommon.Close
		
		
		
		
		
		'******  LogoutMember  ******
		Case "LogoutMember"
			
			
			'Read in username
			strMemberName = Trim(Mid(Request("MemberName"), 1, 20))
			strMemberName = formatSQLInput(strMemberName)
			
			
			'SQL
			strSQL = "SELECT " & strDbTable & "Author.Author_ID, " & strDbTable & "Author.Username, " & strDbTable & "Author.User_code " & _
			"FROM " & strDbTable & "Author" & strDBNoLock & " " & _
			"WHERE " & strDbTable & "Author.Username = '" & strMemberName & "'; "
			
			'Query the database
			rsCommon.Open strSQL, adoCon
			
			'If nothing returned then an error
			If rsCommon.EOF Then
				
				intErrorCode = -150
				strErrorDescription = "Member not found"
			
			'Else member is found so write XML	
			Else
				ReDim Preserve sarryRecords(0)
				
				'Generate new usercode for user
				strMemberCode = userCode(strMemberName)
				
					
				'Update db
				strSQL = "UPDATE " & strDbTable & "Author" & strRowLock & " " & _
				"SET " & _
				strDbTable & "Author.User_code = '" & strMemberCode & "' " & _
				"WHERE " & strDbTable & "Author.Username = '" & strMemberName & "'; "
					
				'Write to the database
				adoCon.Execute(strSQL)
					
					
				sarryRecords(0) = ("" & _
				vbCrLf & "   <Username>" & Server.HTMLEncode(rsCommon("Username")) & "</Username>" & _
				vbCrLf & "   <UserID>" & rsCommon("Author_ID") & "</UserID>" & _
				vbCrLf & "   <MemberCode>" & strMemberCode & "</MemberCode>" & _
				vbCrLf & "   <LoggedOut>True</LoggedOut>")
			
			End If
			
			'Reset Server Objects
			rsCommon.Close
			
			
			
			
		
		
		'******  CreateMember  ******
		Case "CreateNewMember"
			
			
			'Read in username
			strMemberName = Trim(Mid(Request("MemberName"), 1, 20))
			strMemberName = formatSQLInput(strMemberName)
			
			
			
			'******************************************
			'***   Get the starting group ID	***
			'******************************************
	
			'Get the starting group ID number
	
			'Initalise the strSQL variable with an SQL statement to query the database
			strSQL = "SELECT " & strDbTable & "Group.Group_ID " & _
			"FROM " & strDbTable & "Group" & strDBNoLock & " " & _
			"WHERE " & strDbTable & "Group.Starting_group = " & strDBTrue & ";"
	
			'Query the database
			rsCommon.Open strSQL, adoCon
	
			'Get the forum starting group ID number
			intForumStartingGroup = CInt(rsCommon("Group_ID"))
	
			'Close the recordset
			rsCommon.Close
			
			
			
			
			'******************************************
			'***  Read in member details from form	***
			'******************************************
		
		        'Read in the users details from the form
		        strUsername = Trim(Mid(Request("MemberName"), 1, 20))
		        strPassword = Trim(Mid(Request("MemberPassword"), 1, 15))
			strEmail = Trim(Mid(Request("Email"), 1, 60))
			strRealName = Trim(Mid(Request("RealName"), 1, 27))
			strGender = Trim(Mid(Request("Gender"), 1, 10))
			strHomepage = Trim(Mid(Request("Homepage"), 1, 48))
			strSignature = Mid(Request("Signature"), 1, 200)
			If isBool(Request("SignatureAttach")) Then blnAttachSignature = BoolC(Request("SignatureAttach"))  Else blnAttachSignature = True 
			'Check that the ICQ number is a number before reading it in
			If isNumeric(Request("ICQ")) Then strICQNum = Trim(Mid(Request("ICQ"), 1, 15))
			blnShowEmail = False
			blnPMNotify = True
			strDateFormat = Trim(Mid(Request("DateFormat"), 1, 10))
			strTimeOffSet = "+"
			intTimeOffSet = 0
			blnReplyNotify = False
			If isBool(Request("WYSIWYGeditor")) Then blnWYSIWYGEditor = BoolC(Request("WYSIWYGeditor")) Else blnWYSIWYGEditor = True
			If isBool(Request("Active")) Then blnUserActive = BoolC(Request("Active")) Else blnUserActive = True 
		        If isNumeric(Request("GroupID")) Then intUsersGroupID = IntC(Request("GroupID")) Else intUsersGroupID = intForumStartingGroup 
		        If isNumeric(Request("NoOfPosts")) Then lngPosts = LngC(Request("NoOfPosts")) Else lngPosts = 0 
		        If isNumeric(Request("Points")) Then lngPoints = LngC(Request("Points")) Else lngPoints = 0 
		        If isNumeric(Request("Thanked")) Then lngThanked = LngC(Request("Thanked")) Else lngThanked = 0 
		        strMemberTitle = Trim(Mid(Request("MemberTitle"), 1, 40))
		        If isBool(Request("Suspended")) Then blnSuspended = BoolC(Request("Suspended")) Else blnSuspended = False 
		        strAdminNotes = Trim(Mid(removeAllTags(Request("AdminNotes")), 1, 255))
		        If isBool(Request("Newsletter")) Then blnNewsletter = BoolC(Request("Newsletter")) Else blnNewsletter = False 
		
			'If the Group ID is 0 then use the strating group ID instead 
			If intUsersGroupID = 0 Then intUsersGroupID = intForumStartingGroup 
		
		
		        '******************************************
			'***     Read in the avatar		***
			'******************************************
		
		       strAvatar = Trim(Mid(Request("Avatar"), 1, 95))
		
		       'If the avatar is the blank image then the user doesn't want one
		       If strAvatar = strImagePath & "blank.gif" Then strAvatar = ""
		        
		
		
		        '******************************************
			'***     Clean up member details	***
			'******************************************
		
		        'Clean up user input
			strRealName = removeAllTags(strRealName)
			strRealName = formatInput(strRealName)
			strGender = removeAllTags(strGender)
			strGender = formatInput(strGender)
			        
			'Call the function to format the signature
			strSignature = FormatPost(strSignature)
			
			'Call the function to format forum codes
			strSignature = FormatForumCodes(strSignature)
			
			 'Call the filters to remove malcious HTML code
			 strSignature = HTMLsafe(strSignature)
			
			
			'If the user has not entered a hoempage then make sure the homepage variable is blank
			If strHomepage = "http://" Then strHomepage = ""
			
		            
			strMemberTitle = removeAllTags(strMemberTitle) 
			strMemberTitle = formatInput(strMemberTitle)
			
			
		
			'******************************************
			'***     Check the avatar is OK		***
			'******************************************
			'If there is no . in the link then there is no extenison and so can't be an image
		        If inStr(1, strAvatar, ".", 1) = 0 Then
		                  strAvatar = ""
		               
		         'Else remove malicious code and check the extension is an image extension
		         Else
		                'Call the filter for the image
		                strAvatar = formatInput(strAvatar)
		         End If
				
		
			'******************************************
			'*** 	     Create a usercode 		***
			'******************************************
		
		        'Calculate a code for the user
		        strUserCode = userCode(strUsername)
		
			
		
			'******************************************
			'*** 		Encrypt password	***
			'******************************************
		
		        'Encrypt password
			If strPassword <> "" Then
				
				'Encrypt password
				If blnEncryptedPasswords Then																							
			
					'Genrate a slat value
				       	strSalt = getSalt(Len(strPassword))
				
				       'Concatenate salt value to the password
				       strEncryptedPassword = strPassword & strSalt
				
				       'Encrypt the password
				       strEncryptedPassword = HashEncode(strEncryptedPassword)
				
				'Else the password is not set to be encrypted so place the un-encrypted password into the strEncryptedPassword variable
				Else
			
					strEncryptedPassword = strPassword
				End If
			 End If
			
			'******************************************
			'*** 		Date Format	***
			'******************************************
			
			Select Case strDateFormat
					
				'Format dd/mm/yy
				Case "dd/mm/yy"
					strDateFormat = "dd/mm/yy"
					
				'Format mm/dd/yy
				Case "mm/dd/yy"
					strDateFormat = "mm/dd/yy"	
				
				'Format yy/dd/mm
				Case "yy/dd/mm"
					strDateFormat = "yy/dd/mm"
					
				'Format yy/mm/dd
				Case "yy/mm/dd"
					strDateFormat = "yy/mm/dd"
				
				Case Else
					strDateFormat = "dd/mm/yy"		
			
			End Select
			
			
			'SQL
			'Intialise the strSQL variable with an SQL string to open a record set for the Author table
		        strSQL = "SELECT " & strDbTable & "Author.Author_ID, " & strDbTable & "Author.Group_ID, " & strDbTable & "Author.Username, " & strDbTable & "Author.Real_name, " & strDbTable & "Author.Gender, " & strDbTable & "Author.User_code, " & strDbTable & "Author.Password, " & strDbTable & "Author.Salt, " & strDbTable & "Author.Author_email, " & strDbTable & "Author.Homepage, " & strDbTable & "Author.Location, " & strDbTable & "Author.MSN, " & strDbTable & "Author.Yahoo, " & strDbTable & "Author.ICQ, " & strDbTable & "Author.AIM, " & strDbTable & "Author.Occupation, " & strDbTable & "Author.Interests, " & strDbTable & "Author.DOB, " & strDbTable & "Author.Signature, " & strDbTable & "Author.No_of_posts, " & strDbTable & "Author.No_of_PM, " & strDbTable & "Author.Join_date, " & strDbTable & "Author.Avatar, " & strDbTable & "Author.Avatar_title, " & strDbTable & "Author.Last_visit, " & strDbTable & "Author.Time_offset, " & strDbTable & "Author.Time_offset_hours, " & strDbTable & "Author.Date_format, " & strDbTable & "Author.Show_email, " & strDbTable & "Author.Attach_signature, " & strDbTable & "Author.Active, " & strDbTable & "Author.Rich_editor, " & strDbTable & "Author.Reply_notify, " & strDbTable & "Author.PM_notify, " & strDbTable & "Author.Skype, " & strDbTable & "Author.Login_attempt, " & strDbTable & "Author.Points, " & strDbTable & "Author.Thanked, " & strDbTable & "Author.Banned, " & strDbTable & "Author.Info, " & strDbTable & "Author.Newsletter " &_
			"FROM " & strDbTable & "Author" & strRowLock & " " & _
			"WHERE " & strDbTable & "Author.Username = '" & strMemberName & "'; "
		
		        'Set the cursor type property of the record set to Forward Only
		        rsCommon.CursorType = 0
		
		        'Set the Lock Type for the records so that the record set is only locked when it is updated
		        rsCommon.LockType = 3
		
		        'Open the author table
		        rsCommon.Open strSQL, adoCon
				
			
			'If a member is returned then they already exist
			If NOT rsCommon.EOF OR Len(strMemberName) < 2 Then
				
				intErrorCode = -250
				strErrorDescription = "Member already exists"
			
			'If member name less than 3
			ElseIf Len(strMemberName) < 3 Then
				
				intErrorCode = -260
				strErrorDescription = "Member Username to short"
			
			'If password is less than 4
			ElseIf  Len(strPassword) < 4 Then
				
				intErrorCode = -270
				strErrorDescription = "Password to short"
			
			'Else member is found so write XML	
			Else
				ReDim Preserve sarryRecords(0)
				
				With rsCommon
				
					.AddNew
					
					.Fields("Username") = strUsername
	                   		.Fields("Join_date") = internationalDateTime(Now())
					.Fields("Last_visit") = internationalDateTime(Now())
					.Fields("Password") = strEncryptedPassword
				        .Fields("Salt") = strSalt
			                .Fields("User_code") = strUserCode
			                .Fields("Author_email") = strEmail
		                        .Fields("Real_name") = strRealName
		                        .Fields("Gender") = strGender
				       	.Fields("Avatar") = strAvatar
				        .Fields("Homepage") = strHomepage
				        .Fields("Signature") = strSignature
				        .Fields("Attach_signature") = blnAttachSignature
			             	.Fields("Date_format") = strDateFormat
					.Fields("Time_offset") = strTimeOffSet
		 			.Fields("Time_offset_hours") = intTimeOffSet
			    		.Fields("Reply_notify") = blnReplyNotify
			          	.Fields("Rich_editor") = blnWYSIWYGEditor
			          	.Fields("PM_notify") = blnPMNotify
			       		.Fields("Show_email") = blnShowEmail 
		                        .Fields("Newsletter") = blnNewsletter
					.Fields("Group_ID") = intUsersGroupID
					.Fields("Active") = blnUserActive
					.Fields("Banned") = blnSuspended
		                        .Fields("Avatar_title") = strMemberTitle
					.Fields("No_of_posts") = lngPosts
					.Fields("Points") = lngPoints
					.Fields("Thanked") = lngThanked
					.Fields("Info") = strAdminNotes
		                	
		                	
		
		                        'Update the database with the new user's details (needed for MS Access which can be slow updating)
		                        .Update
		
		                        'Re-run the query to read in the updated recordset from the database
		                        .Requery
		                        
		                        
		                        sarryRecords(0) = ("" & _
					vbCrLf & "   <Username>" & Server.HTMLEncode(rsCommon("Username")) & "</Username>" & _
					vbCrLf & "   <UserID>" & rsCommon("Author_ID") & "</UserID>" & _
					vbCrLf & "   <GroupID>" & rsCommon("Group_ID") & "</GroupID>" & _
					vbCrLf & "   <MemberCode>" & rsCommon("User_code") & "</MemberCode>")
					If blnEncryptedPasswords Then	
						sarryRecords(0) = sarryRecords(0) & ("" & _
						vbCrLf & "   <EncryptedPassword>" & rsCommon("Password") & "</EncryptedPassword>" & _
						vbCrLf & "   <Salt>" & rsCommon("Salt") & "</Salt>")
					Else
						sarryRecords(0) = sarryRecords(0) & ("" & _
						vbCrLf & "   <Password>" & rsCommon("Password") & "</Password>")
					End If	
					sarryRecords(0) = sarryRecords(0) & ("" & _
					vbCrLf & "   <Active>" & CBool(rsCommon("Active")) & "</Active>" & _
					vbCrLf & "   <Suspended>" & CBool(rsCommon("Banned"))  & "</Suspended>")
					If isDate(rsCommon("Join_date")) Then sarryRecords(0) = sarryRecords(0) & vbCrLf & "   <Joined>" & internationalDateTime(CDate(rsCommon("Join_date"))) & "</Joined>" Else sarryRecords(0) = sarryRecords(0) & vbCrLf & "   <Joined/>"
					If isDate(rsCommon("Last_visit")) Then sarryRecords(0) = sarryRecords(0) & vbCrLf & "   <LastVisit>" & internationalDateTime(CDate(rsCommon("Last_visit"))) & "</LastVisit>" Else sarryRecords(0) = sarryRecords(0) & vbCrLf & "   <LastVisit/>"
					sarryRecords(0) = sarryRecords(0) & ("" & _
					vbCrLf & "   <Email>" & rsCommon("Author_email") & "</Email>" & _
					vbCrLf & "   <Name>" & Server.HTMLEncode(rsCommon("Real_name")) & "</Name>")
					If isDate(rsCommon("DOB")) Then sarryRecords(0) = sarryRecords(0) & vbCrLf & "   <DOB>" & internationalDateTime(CDate(rsCommon("DOB"))) & "</DOB>" Else sarryRecords(0) = sarryRecords(0) & vbCrLf & "   <DOB/>"
					sarryRecords(0) = sarryRecords(0) & ("" & _
					vbCrLf & "   <Gender>" & Server.HTMLEncode(rsCommon("Gender")) & "</Gender>" & _
					vbCrLf & "   <PostCount>" & rsCommon("No_of_posts") & "</PostCount>" & _
					vbCrLf & "   <Points>" & rsCommon("Points") & "</Points>" & _
					vbCrLf & "   <Thanked>" & rsCommon("Thanked") & "</Thanked>" & _
					vbCrLf & "   <Newsletter>" & CBool(rsCommon("Newsletter")) & "</Newsletter>")
				End with
			
			End If
			
			'Reset Server Objects
			rsCommon.Close
		
		
		
			
		
		'Else no action found
		Case Else
			
			intErrorCode = -400
			strErrorDescription = "Unable to find method '" & strApiAction & "'"
			
	End Select
	
	
	'Close DB
	Call closeDatabase()
	
	
	
	
	
	'******  write XML *******
	
	'If an error has occured display is
	If intErrorCode <> 0 Then
		Response.Write("" & _
		vbCrLf & "<ApiResponse>" & _
		vbCrLf & " <ErrorCode>" & intErrorCode & "</ErrorCode>" & _
		vbCrLf & " <ErrorDescription>" & strErrorDescription & "</ErrorDescription>" & _
		vbCrLf & " <ResultData/>" & _
		vbCrLf & "</ApiResponse>")
		
	
	'Else no error has occured
	Else
		
		Response.Write("" & _
		vbCrLf & "<ApiResponse>" & _
		vbCrLf & " <ErrorCode>0</ErrorCode>" & _
		vbCrLf & " <ErrorDescription/>" & _
		vbCrLf & " <ResultData recordcount=""" & CLng(Ubound(sarryRecords,1) + 1) & """>")
		
		'Loop throug array and display all the records
		For intRecordLoop = 0 TO Ubound(sarryRecords,1)
		
			Response.Write("" & _
			vbCrLf & "  <Record>" & _
			vbCrLf & "   " & sarryRecords(intRecordLoop) & _
			vbCrLf & "  </Record>")
		Next
		
		Response.Write("" & _
		vbCrLf & " </ResultData>" & _
		vbCrLf & "</ApiResponse>")
	End If


'Else there is no XML action so display HTTP XML API documentation
Else
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Web Wiz Forums HTTP XML API Documentation</title>
<meta name="generator" content="Web Wiz Forums" />
<%

Response.Write(vbCrLf  & vbCrLf & "<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Forums(TM) ver. " & strVersion & "" & _
vbCrLf & "Info: http://www.webwizforums.com" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->" & vbCrLf)
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******

%>
<link href="css_styles/default/default_style.css" rel="stylesheet" type="text/css" />
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
<link rel="icon" href="favicon.ico" type="image/x-icon" />
<link rel="shortcut icon" href="favicon.ico" type="image/x-icon" />
</head>
<body>
<img src="<% = strImagePath %>web_wiz_forums.png" alt="Web Wiz Forums Logo" hspace="10" vspace="5" align="middle" />
<h1>&nbsp;&nbsp;Web Wiz Forums HTTP XML API Documentation </h1>
<br />
<table width="98%" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
 <tr>
  <td class="tableLedger"> Version: </td>
  <td class="tableLedger"><% = strAPIversion %></td>
 </tr>
 <tr>
  <td valign="top" class="tableRow"><strong> Description </strong></td>
  <td class="tableRow"> This API has been designed to be used by    external systems. The documentation
   below illustrates the methods that are available in this API. You can
   click on the method name in the list of <a href="#AvailableMethods">Available Methods</a> below to get more details about the parameters that
   it needs and also a test form to try out the method against your own control panel.<br />
   <br />
   The API talks directly to the Web Wiz Forums database and therefore is not effected by any settings or configuration options that may have been set within Web Wiz Forums.</td>
 </tr>
</table>
<br />
<table width="98%" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
 <tr>
  <td class="tableLedger"> Usage</td>
 </tr>
 <tr>
  <td valign="top" class="tableRow">Use
   &quot;HttpAPI.asp?action=[MethodName]&amp;Username=[MasterAdminLoginName]&amp;Password=[MasterAdminPassword]&amp;{AdditionalData}&quot; <br />
   <br />
   e.g.&quot;HttpAPI.asp?action=GetUsers&amp;Username=Adminsitrator&amp;Password=letmein&quot; <br />
   <br />
   Both GET and POST methods are supported and results are returned in XML format.<br />
   <br />
   You can use the Microsoft HTTPXML object in your own ASP or ASP.NET applications to call methods using this HTTP API, the result will be returned in XML format.</td>
 </tr>
</table>
<br />
<br />
<h2><a name="AvailableMethods" id="AvailableMethods"></a>Available Methods</h2>
<table align="center" cellpadding="3" cellspacing="1" class="tableBorder">
 <tbody>
  <tr>
   <td colspan="2" class="tableLedger">General</td>
  </tr>
  <tr>
   <td width="17%" class="tableSubLedger">Name</td>
   <td width="83%" class="tableSubLedger">Description</td>
  </tr>
  <tr>
   <td class="tableRow"><a href="#APIVersion">APIVersion</a></td>
   <td class="tableRow">Returns the current version of this API</td>
  </tr>
  <tr>
   <td class="tableRow"><a href="#WebWizForumsVersion">WebWizForumsVersion</a></td>
   <td class="tableRow">Returns the current version of Web Wiz Forums</td>
  </tr>
 </tbody>
</table>
<br />
<table align="center" cellpadding="3" cellspacing="1" class="tableBorder">
 <tbody>
  <tr>
   <td colspan="2" class="tableLedger">Member Management</td>
  </tr>
  <tr>
   <td width="17%" class="tableSubLedger">Name</td>
   <td width="83%" class="tableSubLedger">Description</td>
  </tr>
  <tr>
   <td class="tableRow"><a href="#CreateNewMember">CreateNewMember</a></td>
   <td class="tableRow">Creates a new Web Wiz Forums Member</td>
  </tr>
  <tr>
   <td class="tableRow"><a href="#GetMemberByName">GetMemberByName</a></td>
   <td class="tableRow">Returns details on the member by Member Name</td>
  </tr>
  <tr>
   <td class="tableRow"><a href="#GetMemberByID">GetMemberByID</a></td>
   <td class="tableRow">Returns details on the member by Member ID (Author_ID in database)</td>
  </tr>
  <tr>
   <td class="tableRow"><a href="#ActivateMember">ActivateMember</a></td>
   <td class="tableRow">Sets members account status to active</td>
  </tr>
  <tr>
   <td class="tableRow"><a href="#SuspendMember">SuspendMember</a></td>
   <td class="tableRow">Sets members account status to suspended</td>
  </tr>
  <tr>
   <td class="tableRow"><a href="#UnsubspendMember">UnsubspendMember</a></td>
   <td class="tableRow">Remove suspended status from members account</td>
  </tr>
  <tr>
   <td class="tableRow"><a href="#LoginMemberCookie">LoginMemberCookie</a></td>
   <td class="tableRow">Returns members login cookie data, to create your own login cookie to the forum</td>
  </tr>
  <tr>
   <td class="tableRow"><a href="#LogoutMember">LogoutMember</a></td>
   <td class="tableRow">Forces logout of member</td>
  </tr>
  <tr>
   <td class="tableRow"><a href="#ChangeMemberPassword">ChangeMemberPassword</a></td>
   <td class="tableRow">Changes the password for a members account</td>
  </tr>
 </tbody>
</table>
<br />
<table align="center" cellpadding="3" cellspacing="1" class="tableBorder">
 <tbody>
  <tr>
   <td colspan="2" class="tableLedger">Forum Management</td>
  </tr>
  <tr>
   <td width="17%" class="tableSubLedger">Name</td>
   <td width="83%" class="tableSubLedger">Description</td>
  </tr>
  <tr>
   <td class="tableRow"><a href="#GetForums">GetForums</a></td>
   <td class="tableRow">Returns  list of forums</td>
  </tr>
  <tr>
   <td class="tableRow"><a href="#LockForumByID">LockForumByID</a></td>
   <td class="tableRow">Locks a Forum by Forum ID (Forum_ID in database) to prevent new posts being made</td>
  </tr>
  <tr>
   <td class="tableRow"><a href="#UnLockForumByID">UnLockForumByID</a></td>
   <td class="tableRow">UnLocks a Forum by Forum ID (Forum_ID in database) </td>
  </tr>
 </tbody>
</table>
<br />
<table align="center" cellpadding="3" cellspacing="1" class="tableBorder">
 <tbody>
  <tr>
   <td colspan="2" class="tableLedger">Topic Management</td>
  </tr>
  <tr>
   <td width="17%" class="tableSubLedger">Name</td>
   <td width="83%" class="tableSubLedger">Description</td>
  </tr>
  <tr>
   <td class="tableRow"><a href="#GetTopicNameByID">GetTopicNameByID</a></td>
   <td class="tableRow">Returns details on the topic by Topic ID (Topic_ID in database)</td>
  </tr>
  <tr>
   <td class="tableRow"><a href="#GetLastPosts">GetLastTopics</a></td>
   <td class="tableRow">Gets the Last Topics (Unlike RSS Feeds this also returns results for non-public forums)</td>
  </tr>
  <tr>
   <td class="tableRow"><a href="#GetLastTopicsByForumID">GetLastTopicsByForumID</a></td>
   <td class="tableRow">Gets the Last Topics by Forum ID (Forum_ID in database)</td>
  </tr>
  <tr>
   <td class="tableRow"><a href="#CloseTopicByID">CloseTopicByID</a></td>
   <td class="tableRow">Closes Topic by ID</td>
  </tr>
  <tr>
   <td class="tableRow"><a href="#OpenTopicByID">OpenTopicByID</a></td>
   <td class="tableRow">Opens Locked  Topic By ID</td>
  </tr>
 </tbody>
</table>
<br />
<table align="center" cellpadding="3" cellspacing="1" class="tableBorder">
 <tbody>
  <tr>
   <td colspan="2" class="tableLedger">Post Management</td>
  </tr>
  <tr>
   <td width="17%" class="tableSubLedger">Name</td>
   <td width="83%" class="tableSubLedger">Description</td>
  </tr>
  <tr>
   <td class="tableRow"><a href="#GetLastPosts">GetLastPosts</a></td>
   <td class="tableRow">Gets the Last Posts  across all forums (Unlike RSS Feeds this also returns results for non-public forums)</td>
  </tr>
  <tr>
   <td class="tableRow"><a href="#GetLastPostsByForumID">GetLastPostsByForumID</a></td>
   <td class="tableRow">Gets the Last Topics  from a particular forum by Forum ID (Forum_ID in database)</td>
  </tr>
 </tbody>
</table>
<br />
<br />
<h2><span class="tableLedger"><a name="Returns" id="Returns"></a></span>Returned Results<br />
</h2>
<table align="center" cellpadding="3" cellspacing="1" class="tableBorder">
 <tbody>
  <tr>
   <td class="tableLedger">Returned Results</td>
  </tr>
  <tr>
   <td class="tableRow"><b><u>Successful Result in XML</u></b><br />
    The following is an example of data being returned if the method is successful. The ErrorCode will be 0 with a blank ErrorDescription. Depending on the method will determine how many records are returned in the ResultData.<br />
    <br />
    <code> &lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot; standalone=&quot;yes&quot; ?&gt;<br />
    &lt;ApiResponse&gt;<br />
    &nbsp;&lt;ErrorCode&gt;0&lt;/ErrorCode&gt;<br />
    &nbsp;&lt;ErrorDescription/&gt;<br />
    &nbsp;&lt;ResultData recordcount=&quot;2&quot;&gt;<br />
    &nbsp;&nbsp;&lt;Record&gt;<br />
    &nbsp;&nbsp;&nbsp;&nbsp;&lt;FirstField&gt;first Field value&lt;/FirstField&gt;<br />
    &nbsp;&nbsp;&nbsp;&nbsp;&lt;SecondField&gt;second Field value&lt;/SecondField&gt;<br />
    &nbsp;&nbsp;&lt;/Record&gt;<br />
    &nbsp;&nbsp;&lt;Record&gt;<br />
    &nbsp;&nbsp;&nbsp;&nbsp;&lt;FirstField&gt;more data&lt;/FirstField&gt;<br />
    &nbsp;&nbsp;&nbsp;&nbsp;&lt;SecondField&gt;more data again&lt;/SecondField&gt;<br />
    &nbsp;&nbsp;&lt;/Record&gt;<br />
    &nbsp;&lt;/ResultData&gt;<br />
    &lt;/ApiResponse&gt;</code><br />
    <br />
    <b><u>Unsuccessful Result in XML</u></b><br />
The following is an example of data being returned if the method is unsuccessful. The ErrorCode will be a negative number (eg. -150), and the ErrorDescription will contain a description of the error. The ResultData will be blank.<br />
<br />
<code> &lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot; standalone=&quot;yes&quot; ?&gt;<br />
&lt;ApiResponse&gt;<br />
&nbsp;&lt;ErrorCode&gt;-150&lt;/ErrorCode&gt;<br />
&nbsp;&lt;ErrorDescription&gt;Admin Login Fail&lt;/ErrorDescription&gt;<br />
&nbsp;&lt;ResultData/&gt;<br />
&lt;/ApiResponse&gt;</code></td>
  </tr>
 </tbody>
</table>
<br />
<br />
<h2>General Method Details<br /></h2>
<h3><a name="APIVersion" id="APIVersion"></a>APIVersion
Returns the current version of this API</h3><br />

<h4>Test Form</h4>
<form action="HttpAPI.asp" method="get" name="test-APIVersion" target="test-APIVersion" id="test-APIVersion">
<input name="action" value="APIVersion" type="hidden" />
<table width="98%" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
  <tr>
   <td width="16%" class="tableSubLedger">Parameter</td>
   <td width="84%" class="tableSubLedger">Value</td>
  </tr>
  <tr>
   <td class="tableRow"> Username* </td>
   <td class="tableRow"><input name="Username" value="" type="text" />
    (required) authentication master admin account username</td>
  </tr>
  <tr>
   <td class="tableRow"> Password* </td>
   <td class="tableRow"><input name="Password" value="" type="password" />
    (required) authentication master admin account password</td>
  </tr>
  <tr>
   <td class="tableRow"></td>
   <td class="tableRow"><input value="Test APIVersion" type="submit" /></td>
  </tr>
</table>
</form>
<br />
Back to <a href="#">Top</a> or <a href="#AvailableMethods">Available Methods</a><br />
<br />
<h3>
<a name="WebWizForumsVersion" id="WebWizForumsVersion"></a>WebWizForumsVersion
</h3>
<p><span class="tableRow">Returns the current version of Web Wiz Forums</span><br />
</p>
<h4>Test Form</h4>
<form action="HttpAPI.asp" method="get" name="test-WebWizForumsVersion" target="test-WebWizForumsVersion" id="test-WebWizForumsVersion">
 <input name="action" value="WebWizForumsVersion" type="hidden" />
 <table width="98%" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
  <tr>
   <td width="16%" class="tableSubLedger">Parameter</td>
   <td width="84%" class="tableSubLedger">Value</td>
  </tr>
  <tr>
   <td class="tableRow"> Username* </td>
   <td class="tableRow"><input name="Username" value="" type="text" />
    (required) authentication master admin account username</td>
  </tr>
  <tr>
   <td class="tableRow"> Password* </td>
   <td class="tableRow"><input name="Password" value="" type="password" />
    (required) authentication master admin account password</td>
  </tr>
  <tr>
   <td class="tableRow"></td>
   <td class="tableRow"><input value="Test WebWizForumsVersion" type="submit" /></td>
  </tr>
 </table>
</form>
<br />
Back to <a href="#">Top</a> or <a href="#AvailableMethods">Available Methods</a><br />
<br />
<br />
<h2>Member Management Method Details </h2>
<h3>
 <a name="CreateNewMember" id="CreateNewMember"></a>CreateNewMember </h3>
<p>Creates a new Web Wiz Forums Member<br />
</p>
<h4>Test Form</h4>
<form action="HttpAPI.asp" method="get" name="test-CreateNewMember" target="test-CreateNewMember" id="test-CreateNewMember">
 <input name="action" value="CreateNewMember" type="hidden" />
 <table width="98%" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
  <tr>
   <td width="16%" class="tableSubLedger">Parameter</td>
   <td width="84%" class="tableSubLedger">Value</td>
  </tr>
  <tr>
   <td class="tableRow"> Username* </td>
   <td class="tableRow"><input name="Username" value="" type="text" />
    (required) authentication master admin account username</td>
  </tr>
  <tr>
   <td class="tableRow"> Password* </td>
   <td class="tableRow"><input name="Password" value="" type="password" />
    (required) authentication master admin account password</td>
  </tr>
  <tr>
   <td class="tableRow"> MemberName* </td>
   <td class="tableRow"><input name="MemberName" value="" type="text" />
    (required) min length 3, max length 20</td>
  </tr>
  <tr>
   <td class="tableRow"> MemberPassword* </td>
   <td class="tableRow"><input name="MemberPassword" value="" type="text" />
   (required) min length 4, max length 15</td>
  </tr>
  <tr>
   <td class="tableRow"> Active </td>
   <td class="tableRow"><input name="Active" value="" type="text" />
   True/False, defaults to True if left blank</td>
  </tr>
  <tr>
   <td class="tableRow"> Suspended</td>
   <td class="tableRow"><input name="Suspended" value="" type="text" />
   True/False, defaults to False if left blank</td>
  </tr>
  <tr>
   <td class="tableRow"> GroupID</td>
   <td class="tableRow"><input name="GroupID" value="" type="text" />
   Group_ID Number, defaults to Starting Group ID if left blank</td>
  </tr>
  <tr>
   <td class="tableRow"> Email</td>
   <td class="tableRow"><input name="Email" value="" type="text" />
    max length 50</td>
  </tr>
  <tr>
   <td class="tableRow"> RealName </td>
   <td class="tableRow"><input name="RealName" value="" type="text" />    
    max length 25</td>
  </tr>
  <tr>
   <td class="tableRow"> Gender </td>
   <td class="tableRow"><input name="Gender" value="" type="text" />
    max length 10</td>
  </tr>
  <tr>
   <td class="tableRow"> Homepage</td>
   <td class="tableRow"><input name="Homepage" value="" type="text" />
     max length 48</td>
  </tr>
 <tr>
   <td class="tableRow"> Avatar</td>
   <td class="tableRow"><input name="Avatar" value="" type="text" />
     max length 95, URL to avatar</td>
  </tr>
   <td class="tableRow"> Signature</td>
   <td class="tableRow"><input name="Signature" value="" type="text" />    
    max length 200, BBcode enabled</td>
  </tr>
  <tr>
   <td class="tableRow"> SignatureAttach</td>
   <td class="tableRow"><input name="SignatureAttach" type="text" />
    True/False, defaults to True if left blank</td>
  </tr>
  <tr>
   <td class="tableRow"> ICQ</td>
   <td class="tableRow"><input name="ICQ" value="" type="text" />
    Number, max length 15</td>
  </tr>
  <tr>
   <td class="tableRow"> DateFormat</td>
   <td class="tableRow"><input name="DateFormat" value="" type="text" />
    Accepted values; dd/mm/yy, mm/dd/yy, yy/dd/mm, yy/mm/dd, defualts to  dd/mm/yy</td>
  </tr>
  <tr>
   <td class="tableRow"> WYSIWYGeditor</td>
   <td class="tableRow"><input name="WYSIWYGeditor" value="" type="text" />
   True/False, defaults to True if left blank</td>
  </tr>
  <tr>
   <td class="tableRow"> NoOfPosts</td>
   <td class="tableRow"><input name="NoOfPosts" value="" type="text" />
    Number, defualts to 0 if left blank</td>
  </tr>
  <tr>
   <td class="tableRow"> Points</td>
   <td class="tableRow"><input name="Points" value="" type="text" />
    Number, defualts to 0 if left blank</td>
  </tr>
  <tr>
   <td class="tableRow"> Thanked</td>
   <td class="tableRow"><input name="Thanked" value="" type="text" />
    Number, defualts to 0 if left blank</td>
  </tr>
  <tr>
   <td class="tableRow"> MemberTitle</td>
   <td class="tableRow"><input name="MemberTitle" value="" type="text" />
    max length 40</td>
  </tr>
  <tr>
   <td class="tableRow"> AdminNotes</td>
   <td class="tableRow"><input name="AdminNotes" value="" type="text" />
    max length 250</td>
  </tr>
  <tr>
   <td class="tableRow">Newsletter</td>
   <td class="tableRow"><input name="Newsletter" value="" type="text" />
   True/False, defaults to False if left blank</td>
  </tr>
  <tr>
   <td class="tableRow"></td>
   <td class="tableRow"><input value="Test CreateNewMember" type="submit" /></td>
  </tr>
 </table>
</form>
<br />
Back to <a href="#">Top</a> or <a href="#AvailableMethods">Available Methods</a>
<h3><a name="GetMemberByName" id="GetMemberByName"></a>GetMemberByName </h3>
<p>Returns details on the member<br />
</p>
<h4>Test Form</h4>
<form action="HttpAPI.asp" method="get" name="test-GetMemberByName" target="test-GetMemberByName" id="test-GetMemberByName">
 <input name="action" value="GetMemberByName" type="hidden" />
 <table width="98%" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
  <tr>
   <td width="16%" class="tableSubLedger">Parameter</td>
   <td width="84%" class="tableSubLedger">Value</td>
  </tr>
  <tr>
   <td class="tableRow"> Username* </td>
   <td class="tableRow"><input name="Username" value="" type="text" />
    (required) authentication master admin account username</td>
  </tr>
  <tr>
   <td class="tableRow"> Password* </td>
   <td class="tableRow"><input name="Password" value="" type="password" />
    (required) authentication master admin account password</td>
  </tr>
   <tr>
   <td class="tableRow"> MemberName* </td>
   <td class="tableRow"><input name="MemberName" value="" type="text" />
    (required) </td>
  </tr>
  <tr>
   <td class="tableRow"></td>
   <td class="tableRow"><input value="Test GetMemberByName" type="submit" /></td>
  </tr>
 </table>
</form>
<br />
Back to <a href="#">Top</a> or <a href="#AvailableMethods">Available Methods</a><br />
<br />
<h3>
<a name="GetMemberByID" id="GetMemberByID"></a>GetMemberByID
</h3>
<p>Returns details on the membefrom their member ID (Author_ID)<br />
</p>
<h4>Test Form</h4>
<form action="HttpAPI.asp" method="get" name="test-GetMemberByID" target="test-GetMemberByID" id="test-GetMemberByID">
 <input name="action" value="GetMemberByID" type="hidden" />
 <table width="98%" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
  <tr>
   <td width="16%" class="tableSubLedger">Parameter</td>
   <td width="84%" class="tableSubLedger">Value</td>
  </tr>
  <tr>
   <td class="tableRow"> Username* </td>
   <td class="tableRow"><input name="Username" value="" type="text" />
    (required) authentication master admin account username</td>
  </tr>
  <tr>
   <td class="tableRow"> Password* </td>
   <td class="tableRow"><input name="Password" value="" type="password" />
    (required) authentication master admin account password</td>
  </tr>
  <tr>
   <td class="tableRow"> MemberID* </td>
   <td class="tableRow"><input name="MemberID" value="" type="text" />
    (required) </td>
  </tr>
  <tr>
   <td class="tableRow"></td>
   <td class="tableRow"><input value="Test GetMemberByID" type="submit" /></td>
  </tr>
 </table>
</form>
<br />
Back to <a href="#">Top</a> or <a href="#AvailableMethods">Available Methods</a><br />
<br />
<h3>
<a name="ActivateMember" id="ActivateMember"></a>ActivateMember
</h3>
<p>Sets members account status to active<br />
</p>
<h4>Test Form</h4>
<form action="HttpAPI.asp" method="get" name="test-ActivateMember" target="test-ActivateMember" id="test-ActivateMember">
 <input name="action" value="ActivateMember" type="hidden" />
 <table width="98%" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
  <tr>
   <td width="16%" class="tableSubLedger">Parameter</td>
   <td width="84%" class="tableSubLedger">Value</td>
  </tr>
  <tr>
   <td class="tableRow"> Username* </td>
   <td class="tableRow"><input name="Username" value="" type="text" />
    (required) authentication master admin account username</td>
  </tr>
  <tr>
   <td class="tableRow"> Password* </td>
   <td class="tableRow"><input name="Password" value="" type="password" />
    (required) authentication master admin account password</td>
  </tr>
  <tr>
   <td class="tableRow"> MemberName* </td>
   <td class="tableRow"><input name="MemberName" value="" type="text" />
    (required) </td>
  </tr>
  <tr>
   <td class="tableRow"></td>
   <td class="tableRow"><input value="Test ActivateMember" type="submit" /></td>
  </tr>
 </table>
</form>
<br />
Back to <a href="#">Top</a> or <a href="#AvailableMethods">Available Methods</a><br />
<br />
<h3>
<a name="SuspendMember" id="SuspendMember"></a>SuspendMember
</h3>
<p>Sets members account status to suspended<br />
</p>
<h4>Test Form</h4>
<form action="HttpAPI.asp" method="get" name="test-SuspendMember" target="test-SuspendMember" id="test-SuspendMember">
 <input name="action" value="SuspendMember" type="hidden" />
 <table width="98%" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
  <tr>
   <td width="16%" class="tableSubLedger">Parameter</td>
   <td width="84%" class="tableSubLedger">Value</td>
  </tr>
  <tr>
   <td class="tableRow"> Username* </td>
   <td class="tableRow"><input name="Username" value="" type="text" />
    (required) authentication master admin account username</td>
  </tr>
  <tr>
   <td class="tableRow"> Password* </td>
   <td class="tableRow"><input name="Password" value="" type="password" />
    (required) authentication master admin account password</td>
  </tr>
  <tr>
   <td class="tableRow"> MemberName* </td>
   <td class="tableRow"><input name="MemberName" value="" type="text" />
    (required) </td>
  </tr>
  <tr>
   <td class="tableRow"></td>
   <td class="tableRow"><input value="Test SuspendMember" type="submit" /></td>
  </tr>
 </table>
</form>
<br />
Back to <a href="#">Top</a> or <a href="#AvailableMethods">Available Methods</a><br />
<br />
<h3>
<a name="UnsubspendMember" id="UnsubspendMember"></a>UnsubspendMember
</h3>
<p>Remove suspended status from members account<br />
</p>
<h4>Test Form</h4>
<form action="HttpAPI.asp" method="get" name="test-UnsubspendMember" target="test-UnsubspendMember" id="test-UnsubspendMember">
 <input name="action" value="UnsubspendMember" type="hidden" />
 <table width="98%" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
  <tr>
   <td width="16%" class="tableSubLedger">Parameter</td>
   <td width="84%" class="tableSubLedger">Value</td>
  </tr>
  <tr>
   <td class="tableRow"> Username* </td>
   <td class="tableRow"><input name="Username" value="" type="text" />
    (required) authentication master admin account username</td>
  </tr>
  <tr>
   <td class="tableRow"> Password* </td>
   <td class="tableRow"><input name="Password" value="" type="password" />
    (required) authentication master admin account password</td>
  </tr>
  <tr>
   <td class="tableRow"> MemberName* </td>
   <td class="tableRow"><input name="MemberName" value="" type="text" />
    (required) </td>
  </tr>
  <tr>
   <td class="tableRow"></td>
   <td class="tableRow"><input value="Test UnsubspendMember" type="submit" /></td>
  </tr>
 </table>
</form>
<br />
Back to <a href="#">Top</a> or <a href="#AvailableMethods">Available Methods</a><br />
<br />
<h3> <a name="LoginMemberCookie" id="LoginMemberCookie"></a>LoginMemberCookie </h3>
<p>Returns members login cookie data, to create your own login cookie to the forum<br />
</p>
<h4>Test Form</h4>
<form action="HttpAPI.asp" method="get" name="test-LoginMemberCookie" target="test-LoginMemberCookie" id="test-LoginMemberCookie">
 <input name="action" value="LoginMemberCookie" type="hidden" />
 <table width="98%" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
  <tr>
   <td width="16%" class="tableSubLedger">Parameter</td>
   <td width="84%" class="tableSubLedger">Value</td>
  </tr>
  <tr>
   <td class="tableRow"> Username* </td>
   <td class="tableRow"><input name="Username" value="" type="text" />
    (required) authentication master admin account username</td>
  </tr>
  <tr>
   <td class="tableRow"> Password* </td>
   <td class="tableRow"><input name="Password" value="" type="password" />
    (required) authentication master admin account password</td>
  </tr>
  <tr>
   <td class="tableRow"> MemberName* </td>
   <td class="tableRow"><input name="MemberName" value="" type="text" />
    (required) </td>
  </tr>
  <tr>
   <td class="tableRow"> MemberPassword</td>
   <td class="tableRow"><input name="MemberPassword" value="" type="text" />
    Not required. Use if you want to create a Username and Password login system</td>
  </tr>
  <tr>
   <td class="tableRow"></td>
   <td class="tableRow"><input value="Test LoginMemberCookie" type="submit" /></td>
  </tr>
 </table>
</form>
<br />
Back to <a href="#">Top</a> or <a href="#AvailableMethods">Available Methods</a><br />
<br />
<h3> <a name="LogoutMember" id="LogoutMember"></a>LogoutMember </h3>
<p>Forces logout of member<br />
</p>
<h4>Test Form</h4>
<form action="HttpAPI.asp" method="get" name="test-LogoutMember" target="test-LogoutMember" id="test-LogoutMember">
 <input name="action" value="LogoutMember" type="hidden" />
 <table width="98%" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
  <tr>
   <td width="16%" class="tableSubLedger">Parameter</td>
   <td width="84%" class="tableSubLedger">Value</td>
  </tr>
  <tr>
   <td class="tableRow"> Username* </td>
   <td class="tableRow"><input name="Username" value="" type="text" />
    (required) authentication master admin account username</td>
  </tr>
  <tr>
   <td class="tableRow"> Password* </td>
   <td class="tableRow"><input name="Password" value="" type="password" />
    (required) authentication master admin account password</td>
  </tr>
  <tr>
   <td class="tableRow"> MemberName* </td>
   <td class="tableRow"><input name="MemberName" value="" type="text" />
    (required) </td>
  </tr>
  <tr>
   <td class="tableRow"></td>
   <td class="tableRow"><input value="Test LogoutMember" type="submit" /></td>
  </tr>
 </table>
</form>
<br />
Back to <a href="#">Top</a> or <a href="#AvailableMethods">Available Methods</a><br />
<br />
<h3><a name="ChangeMemberPassword" id="ChangeMemberPassword"></a>ChangeMemberPassword </h3>
<p>Changes the password for a members account<br />
</p>
<h4>Test Form</h4>
<form action="HttpAPI.asp" method="get" name="test-ChangeMemberPassword" target="test-ChangeMemberPassword" id="test-ChangeMemberPassword">
 <input name="action" value="ChangeMemberPassword" type="hidden" />
 <table width="98%" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
  <tr>
   <td width="16%" class="tableSubLedger">Parameter</td>
   <td width="84%" class="tableSubLedger">Value</td>
  </tr>
  <tr>
   <td class="tableRow"> Username* </td>
   <td class="tableRow"><input name="Username" value="" type="text" />
    (required) authentication master admin account username</td>
  </tr>
  <tr>
   <td class="tableRow"> Password* </td>
   <td class="tableRow"><input name="Password" value="" type="password" />
    (required) authentication master admin account password</td>
  </tr>
  <tr>
   <td class="tableRow"> MemberName* </td>
   <td class="tableRow"><input name="MemberName" value="" type="text" />
    (required) </td>
  </tr>
  <tr>
   <td class="tableRow"> NewPassword* </td>
   <td class="tableRow"><input name="NewPassword" value="" type="text" />
    (required) max length 15</td>
  </tr>
  <tr>
   <td class="tableRow"></td>
   <td class="tableRow"><input value="Test ChangeMemberPassword" type="submit" /></td>
  </tr>
 </table>
</form>
<br />
Back to <a href="#">Top</a> or <a href="#AvailableMethods">Available Methods</a><br />
<br />
<br />
<h2>Forum Management Method Details </h2>
<h3>
<a name="GetForums" id="GetForums"></a>GetForums
</h3>
<p>Returns  list of forums<br />
</p>
<h4>Test Form</h4>
<form action="HttpAPI.asp" method="get" name="test-GetForums" target="test-GetForums" id="test-GetForums">
 <input name="action" value="GetForums" type="hidden" />
 <table width="98%" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
  <tr>
   <td width="16%" class="tableSubLedger">Parameter</td>
   <td width="84%" class="tableSubLedger">Value</td>
  </tr>
  <tr>
   <td class="tableRow"> Username* </td>
   <td class="tableRow"><input name="Username" value="" type="text" />
    (required) authentication master admin account username</td>
  </tr>
  <tr>
   <td class="tableRow"> Password* </td>
   <td class="tableRow"><input name="Password" value="" type="password" />
    (required) authentication master admin account password</td>
  </tr>
  <tr>
   <td class="tableRow"></td>
   <td class="tableRow"><input value="Test GetForums" type="submit" /></td>
  </tr>
 </table>
</form>
<br />
Back to <a href="#">Top</a> or <a href="#AvailableMethods">Available Methods</a><br />
<br />
<h3>
<a name="LockForumByID" id="LockForumByID"></a>LockForumByID
</h3>
<p>Locks a Forum by Forum ID (Forum_ID in database) to prevent new posts being made<br />
</p>
<h4>Test Form</h4>
<form action="HttpAPI.asp" method="get" name="test-LockForumByID" target="test-LockForumByID" id="test-LockForumByID">
 <input name="action" value="LockForumByID" type="hidden" />
 <table width="98%" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
  <tr>
   <td width="16%" class="tableSubLedger">Parameter</td>
   <td width="84%" class="tableSubLedger">Value</td>
  </tr>
  <tr>
   <td class="tableRow"> Username* </td>
   <td class="tableRow"><input name="Username" value="" type="text" />
    (required) authentication master admin account username</td>
  </tr>
  <tr>
   <td class="tableRow"> Password* </td>
   <td class="tableRow"><input name="Password" value="" type="password" />
    (required) authentication master admin account password</td>
  </tr>
  <tr>
   <td class="tableRow"> ForumID* </td>
   <td class="tableRow"><input name="ForumID" value="" type="text" />
    (required) </td>
  </tr>
  <tr>
   <td class="tableRow"></td>
   <td class="tableRow"><input value="Test LockForumByID" type="submit" /></td>
  </tr>
 </table>
</form>
<br />
Back to <a href="#">Top</a> or <a href="#AvailableMethods">Available Methods</a><br />
<br />
<h3>
<a name="UnLockForumByID" id="UnLockForumByID"></a>UnLockForumByID
</h3>
UnLocks a Forum by Forum ID (Forum_ID in database)<br />

<h4>Test Form</h4>
<form action="HttpAPI.asp" method="get" name="test-UnLockForumByID" target="test-UnLockForumByID" id="test-UnLockForumByID">
 <input name="action" value="UnLockForumByID" type="hidden" />
 <table width="98%" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
  <tr>
   <td width="16%" class="tableSubLedger">Parameter</td>
   <td width="84%" class="tableSubLedger">Value</td>
  </tr>
  <tr>
   <td class="tableRow"> Username* </td>
   <td class="tableRow"><input name="Username" value="" type="text" />
    (required) authentication master admin account username</td>
  </tr>
  <tr>
   <td class="tableRow"> Password* </td>
   <td class="tableRow"><input name="Password" value="" type="password" />
    (required) authentication master admin account password</td>
  </tr>
  <tr>
   <td class="tableRow"> ForumID* </td>
   <td class="tableRow"><input name="ForumID" value="" type="text" />
    (required) </td>
  </tr>
  <tr>
   <td class="tableRow"></td>
   <td class="tableRow"><input value="Test UnLockForumByID" type="submit" /></td>
  </tr>
 </table>
</form>
<br />
Back to <a href="#">Top</a> or <a href="#AvailableMethods">Available Methods</a><br />
<br />
<br />
<h2>Topic Management Method Details </h2>
<h3> <a name="GetTopicNameByID" id="GetTopicNameByID"></a>GetTopicNameByID </h3>
<p>Returns details on the topic by Topic ID (Topic_ID in database)<br />
</p>
<h4>Test Form</h4>
<form action="HttpAPI.asp" method="get" name="test-GetTopicNameByID" target="test-GetTopicNameByID" id="test-GetTopicNameByID">
 <input name="action" value="GetTopicNameByID" type="hidden" />
 <table width="98%" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
  <tr>
   <td width="16%" class="tableSubLedger">Parameter</td>
   <td width="84%" class="tableSubLedger">Value</td>
  </tr>
  <tr>
   <td class="tableRow"> Username* </td>
   <td class="tableRow"><input name="Username" value="" type="text" />
    (required) authentication master admin account username</td>
  </tr>
  <tr>
   <td class="tableRow"> Password* </td>
   <td class="tableRow"><input name="Password" value="" type="password" />
    (required) authentication master admin account password</td>
  </tr>
  <tr>
   <td class="tableRow"> TopicID* </td>
   <td class="tableRow"><input name="TopicID" value="" type="text" />
    (required) </td>
  </tr>
  <tr>
   <td class="tableRow"></td>
   <td class="tableRow"><input value="Test GetTopicNameByID" type="submit" /></td>
  </tr>
 </table>
</form>
<br />
Back to <a href="#">Top</a> or <a href="#AvailableMethods">Available Methods</a><br />
<h3><a name="GetLastTopics" id="GetLastTopics"></a>GetLastTopics </h3>
<p>Gets the Last Topics  across all forums (Unlike RSS Feeds this also returns results for non-public forums)<br />
</p>
<h4>Test Form</h4>
<form action="HttpAPI.asp" method="get" name="test-GetLastTopics" target="test-GetLastTopics" id="test-GetLastTopics">
 <input name="action" value="GetLastTopics" type="hidden" />
 <table width="98%" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
  <tr>
   <td width="16%" class="tableSubLedger">Parameter</td>
   <td width="84%" class="tableSubLedger">Value</td>
  </tr>
  <tr>
   <td class="tableRow"> Username* </td>
   <td class="tableRow"><input name="Username" value="" type="text" />
    (required) authentication master admin account username</td>
  </tr>
  <tr>
   <td class="tableRow"> Password* </td>
   <td class="tableRow"><input name="Password" value="" type="password" />
    (required) authentication master admin account password</td>
  </tr>
  <tr>
   <td class="tableRow"> MaxResults</td>
   <td class="tableRow"><input name="MaxResults" value="" type="text" />
    Maximum Number of Results to return (max.50)</td>
  </tr>
  <tr>
   <td class="tableRow"></td>
   <td class="tableRow"><input value="Test GetLastTopics" type="submit" /></td>
  </tr>
 </table>
</form>
<br />
Back to <a href="#">Top</a> or <a href="#AvailableMethods">Available Methods</a><br />
<br />
<h3><a name="GetLastTopicsByForumID" id="GetLastTopicsByForumID"></a>GetLastTopicsByForumID </h3>
<p>Gets the Last Topics  from a particular forum by Forum ID (Forum_ID in database)<br />
</p>
<h4>Test Form</h4>
<form action="HttpAPI.asp" method="get" name="test-GetLastTopicsByForumID" target="test-GetLastTopicsByForumID" id="test-GetLastTopicsByForumID">
 <input name="action" value="GetLastTopicsByForumID" type="hidden" />
 <table width="98%" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
  <tr>
   <td width="16%" class="tableSubLedger">Parameter</td>
   <td width="84%" class="tableSubLedger">Value</td>
  </tr>
  <tr>
   <td class="tableRow"> Username* </td>
   <td class="tableRow"><input name="Username" value="" type="text" />
    (required) authentication master admin account username</td>
  </tr>
  <tr>
   <td class="tableRow"> Password* </td>
   <td class="tableRow"><input name="Password" value="" type="password" />
    (required) authentication master admin account password</td>
  </tr>
  <tr>
   <td class="tableRow"> ForumID* </td>
   <td class="tableRow"><input name="ForumID" value="" type="text" />
    (required) </td>
  </tr>
  <tr>
   <td class="tableRow"> MaxResults</td>
   <td class="tableRow"><input name="MaxResults" value="" type="text" />
    Maximum Number of Results to return (max.50)</td>
  </tr>
  <tr>
   <td class="tableRow"></td>
   <td class="tableRow"><input value="Test GetLastTopicsByForumID" type="submit" /></td>
  </tr>
 </table>
</form>
<br />
Back to <a href="#">Top</a> or <a href="#AvailableMethods">Available Methods</a><br />
<br />
<h3> <a name="CloseTopicByID" id="CloseTopicByID"></a>CloseTopicByID </h3>
<p>Closes Topic by Topic ID<br />
</p>
<h4>Test Form</h4>
<form action="HttpAPI.asp" method="get" name="test-CloseTopicByID" target="test-CloseTopicByID" id="test-CloseTopicByID">
 <input name="action" value="CloseTopicByID" type="hidden" />
 <table width="98%" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
  <tr>
   <td width="16%" class="tableSubLedger">Parameter</td>
   <td width="84%" class="tableSubLedger">Value</td>
  </tr>
  <tr>
   <td class="tableRow"> Username* </td>
   <td class="tableRow"><input name="Username" value="" type="text" />
    (required) authentication master admin account username</td>
  </tr>
  <tr>
   <td class="tableRow"> Password* </td>
   <td class="tableRow"><input name="Password" value="" type="password" />
    (required) authentication master admin account password</td>
  </tr>
  <tr>
   <td class="tableRow"> TopicID* </td>
   <td class="tableRow"><input name="TopicID" value="" type="text" />
    (required) </td>
  </tr>
  <tr>
   <td class="tableRow"></td>
   <td class="tableRow"><input value="Test CloseTopicByID" type="submit" /></td>
  </tr>
 </table>
</form>
<br />
Back to <a href="#">Top</a> or <a href="#AvailableMethods">Available Methods</a><br />
<br />
<h3> <a name="OpenTopicByID" id="OpenTopicByID"></a>OpenTopicByID </h3>
<p>Opens Locked  Topic By ID<br />
</p>
<h4>Test Form</h4>
<form action="HttpAPI.asp" method="get" name="test-OpenTopicByID" target="test-OpenTopicByID" id="test-OpenTopicByID">
 <input name="action" value="OpenTopicByID" type="hidden" />
 <table width="98%" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
  <tr>
   <td width="16%" class="tableSubLedger">Parameter</td>
   <td width="84%" class="tableSubLedger">Value</td>
  </tr>
  <tr>
   <td class="tableRow"> Username* </td>
   <td class="tableRow"><input name="Username" value="" type="text" />
    (required) authentication master admin account username</td>
  </tr>
  <tr>
   <td class="tableRow"> Password* </td>
   <td class="tableRow"><input name="Password" value="" type="password" />
    (required) authentication master admin account password</td>
  </tr>
  <tr>
   <td class="tableRow"> TopicID* </td>
   <td class="tableRow"><input name="TopicID" value="" type="text" />
    (required) </td>
  </tr>
  <tr>
   <td class="tableRow"></td>
   <td class="tableRow"><input value="Test OpenTopicByID" type="submit" /></td>
  </tr>
 </table>
</form>
<br />
Back to <a href="#">Top</a> or <a href="#AvailableMethods">Available Methods</a><br />
<br />
<h2>Post Management Method Details<br />
</h2>
<h3><a name="GetLastPosts" id="GetLastPosts"></a>GetLastPosts </h3>
<p>Gets the Last Posts  across all forums (Unlike RSS Feeds this also returns results for non-public forums)<br />
</p>
<h4>Test Form</h4>
<form action="HttpAPI.asp" method="get" name="test-GetLastPosts" target="test-GetLastPosts" id="test-GetLastPosts">
 <input name="action" value="GetLastPosts" type="hidden" />
 <table width="98%" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
  <tr>
   <td width="16%" class="tableSubLedger">Parameter</td>
   <td width="84%" class="tableSubLedger">Value</td>
  </tr>
  <tr>
   <td class="tableRow"> Username* </td>
   <td class="tableRow"><input name="Username" value="" type="text" />
    (required) authentication master admin account username</td>
  </tr>
  <tr>
   <td class="tableRow"> Password* </td>
   <td class="tableRow"><input name="Password" value="" type="password" />
    (required) authentication master admin account password</td>
  </tr>
  <tr>
   <td class="tableRow"> MaxResults</td>
   <td class="tableRow"><input name="MaxResults" value="" type="text" />
   Maximum Number of Results to return (max.50)</td>
  </tr>
  <tr>
   <td class="tableRow"></td>
   <td class="tableRow"><input value="Test GetLastPosts" type="submit" /></td>
  </tr>
 </table>
</form>
<br />
Back to <a href="#">Top</a> or <a href="#AvailableMethods">Available Methods</a><br />
<br />
<h3><a name="GetLastPostsByForumID" id="GetLastPostsByForumID"></a>GetLastPostsByForumID </h3>
<p>Gets the Lasts Posts from a particular forum by Forum ID (Forum_ID in database)<br />
</p>
<h4>Test Form</h4>
<form action="HttpAPI.asp" method="get" name="test-GetLastPostsByForumID" target="test-GetLastPostsByForumID" id="test-GetLastPostsByForumID">
 <input name="action" value="GetLastPostsByForumID" type="hidden" />
 <table width="98%" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
  <tr>
   <td width="16%" class="tableSubLedger">Parameter</td>
   <td width="84%" class="tableSubLedger">Value</td>
  </tr>
  <tr>
   <td class="tableRow"> Username* </td>
   <td class="tableRow"><input name="Username" value="" type="text" />
    (required) authentication master admin account username</td>
  </tr>
  <tr>
   <td class="tableRow"> Password* </td>
   <td class="tableRow"><input name="Password" value="" type="password" />
    (required) authentication master admin account password</td>
  </tr>
   <tr>
   <td class="tableRow"> ForumID* </td>
   <td class="tableRow"><input name="ForumID" value="" type="text" />
    (required) </td>
  </tr>
  <tr>
   <td class="tableRow"> MaxResults</td>
   <td class="tableRow"><input name="MaxResults" value="" type="text" />
    Maximum Number of Results to return (max.50)</td>
  </tr>
  <tr>
   <td class="tableRow"></td>
   <td class="tableRow"><input value="Test GetLastPostsByForumID" type="submit" /></td>
  </tr>
 </table>
</form>
<br />
Back to <a href="#">Top</a> or <a href="#AvailableMethods">Available Methods</a><br />
<br />
<br />
<p>* Field is required when executing the method</p>
</body>
</html>
<%

End If

%>