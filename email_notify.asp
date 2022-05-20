<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
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






'Set the response buffer to true as we maybe redirecting
Response.Buffer = True 

'Declare variables
Dim lngTopicID		'Holds the topic ID
Dim strReturnValue	'Holds the return value of the page
Dim strMode		'Holds the mode of the page
Dim strReturnPage	'Holds the return page
Dim intReturnPageNumber



'Read in the forum or topic ID
intForumID = IntC(Request("FID"))
lngTopicID = LngC(Request.QueryString("TID"))
strMode = Request.QueryString("M")
If isNumeric(Request("PN")) Then intReturnPageNumber = IntC(Request("PN")) Else intReturnPageNumber = 1


'If there is no Forum ID read it in from the session
If intForumID = 0 Then
	intForumID = intC(getSessionItem("FID"))
End If



'If there is no forum ID or Topic ID then send to the main forum page
If intForumID = 0 AND lngTopicID = 0 Then 
	Call closeDatabase()
	Response.Redirect("default.asp" & strQsSID1)
End If






'If this is a Topic to watch then watch or unwatch this topic
If lngTopicID AND blnEmail AND intGroupID <> 2 AND strMode = "" Then Call WatchUnWatchTopic(lngTopicID)

'If this is a Topic to watch then watch or unwatch this topic
If intForumID AND blnEmail AND intGroupID <> 2 AND strMode = "" Then Call WatchUnWatchForum(intForumID)

'If this is a link form an unsubscribe email notify link in an email unwatch this topic or forum
If strMode = "Unsubscribe" AND intForumID <> "" AND lngTopicID <> "" Then Call UnsubscribeEmailNotify(intForumID, lngTopicID)

'If this is from the subscription page then add to the forum watch list
If intForumID AND blnEmail AND intGroupID <> 2 AND strMode = "SP" Then Call WatchUnWatchForum(intForumID)




'******************************************
'***  	  Watch or Unwatch Topic        ***
'******************************************

Private Function WatchUnWatchTopic(lngTopicID)

	'Check the session ID to stop CSRF
	Call checkFormID(Request.QueryString("XID"))
	
	
	'Initalise the SQL string with a query to get the email notify topic details
	strSQL = "SELECT " & strDbTable & "EmailNotify.*  " & _
	"FROM " & strDbTable & "EmailNotify" & strRowLock & " " & _
	"WHERE " & strDbTable & "EmailNotify.Author_ID = " & lngLoggedInUserID & " AND " & strDbTable & "EmailNotify.Topic_ID = " & lngTopicID & ";"

	With rsCommon

		'Set the cursor type property of the record set to Forward Only
		.CursorType = 0

		'Set the Lock Type for the records so that the record set is only locked when it is updated
		.LockType = 3
		
		'Query the database
		.Open strSQL, adoCon


		'If the user no-longer wants email notification for this topic then remove the entry form the db
		If NOT .EOF Then
			
			Do while NOT .EOF

				'Delete the db entry
				.Delete
				
				'Move to next record
				.MoveNext
				
				'Set the return value
				strReturnValue = "&EN=TU"
			Loop

		'Else if this is a new post and the user wants to be notified add the new entry to the database
		Else

			'Check to see if the user is allowed to view posts in this forum
			Call forumPermissions(intForumID, intGroupID)
			
			'If the user can read in this forum the add them
			If blnRead Then
				'Add new rs
				.AddNew
	
				'Create new entry
				.Fields("Author_ID") = lngLoggedInUserID
				.Fields("Topic_ID") = lngTopicID
	
				'Upade db with new rs
				.Update
				
				'Set the return value
				strReturnValue = "&EN=TS"
			End If
		End If

		'Clean up
		.Close

	End With
	
	'Clean up
	Call closeDatabase()
	
	'Return to Topic Page
	Response.Redirect("forum_posts.asp?TID=" & lngTopicID & "&PN=" & intReturnPageNumber & strReturnValue & strQsSID3)
End Function






'******************************************
'***  	  Watch or Unwatch Forum        ***
'******************************************

Private Function WatchUnWatchForum(intForumID)

	'Check the session ID to stop CSRF
	Call checkFormID(Request.QueryString("XID"))

	'Initalise the SQL string with a query to get the email notify forum details
	strSQL = "SELECT " & strDbTable & "EmailNotify.*  " & _
	"FROM " & strDbTable & "EmailNotify" & strRowLock & " " & _
	"WHERE " & strDbTable & "EmailNotify.Author_ID = " & lngLoggedInUserID & " AND " & strDbTable & "EmailNotify.Forum_ID = " & intForumID & ";"

	With rsCommon

		'Set the cursor type property of the record set to Forward Only
		.CursorType = 0

		'Set the Lock Type for the records so that the record set is only locked when it is updated
		.LockType = 3
		
		'Query the database
		.Open strSQL, adoCon


		'If the user no-longer wants email notification for this forum then remove the entry form the db
		If NOT .EOF Then

			'If this is not from teh subscription page then delete
			If strMode <> "SP" Then 
				Do while NOT .EOF
					'Delete the db entry
					.Delete
					
					'Move to next record
					.MoveNext
					
					'Set the return value
					strReturnValue = "&EN=FU"
				Loop
			End If

		'Else if this is a new post and the user wants to be notified add the new entry to the database
		Else

			'Check to see if the user is allowed to view posts in this forum
			Call forumPermissions(intForumID, intGroupID)
			
			'If the user can read in this forum the add them
			If blnRead Then
				
				'Add new rs
				.AddNew
	
				'Create new entry
				.Fields("Author_ID") = lngLoggedInUserID
				.Fields("Forum_ID") = intForumID
	
				'Upade db with new rs
				.Update
				
				'Set the return value
				strReturnValue = "&EN=FS"
			End If
		End If

		'Clean up
		.Close

	End With
	
	'Clean up
	Call closeDatabase()
	
	'Return to Forum Page
	If strMode = "SP" Then
		Response.Redirect("email_notify_subscriptions.asp" & strQsSID1)
	Else
		Response.Redirect("forum_topics.asp?FID=" & intForumID & "&PN=" & intReturnPageNumber & strReturnValue & strQsSID3)
	End If
End Function






'******************************************
'*** Unsubscribe from email notify link ***
'******************************************

Private Function UnsubscribeEmailNotify(intForumID, lngTopicID)


	'If the user is not logged in then send them to the login page
	If intGroupID = 2 Then Response.Redirect("login_user.asp?returnURL=email_notify.asp?FID=" & intForumID & "%26TID=" & lngTopicID & "%26M=Unsubscribe" & strQsSID3)
	
	'Initalise the SQL string with a query to get the email notify topic details
	strSQL = "SELECT " & strDbTable & "EmailNotify.*  " & _
	"FROM " & strDbTable & "EmailNotify" & strRowLock & " " & _
	"WHERE " & strDbTable & "EmailNotify.Author_ID = " & lngLoggedInUserID & " AND " & strDbTable & "EmailNotify.Topic_ID = " & lngTopicID & ";"

	With rsCommon

		'Set the cursor type property of the record set to Forward Only
		.CursorType = 0

		'Set the Lock Type for the records so that the record set is only locked when it is updated
		.LockType = 3
		
		'Query the database
		.Open strSQL, adoCon


		'If a record is returned then the user is subscribed to the topic so delete their email notification
		If NOT .EOF Then

			Do while NOT .EOF
				'Delete the db entry
				.Delete
				
				'Move to next record
				.MoveNext
				
				'Set the return value
				strReturnValue = "&EN=TU"
				strReturnPage = "forum_posts.asp?TID=" & lngTopicID & strReturnValue
			Loop

		'Else the user is probally got forum post notification so check the db and delete that if they do
		Else
			'Clean up
			.Close
			
			'Initalise the SQL string with a query to get the poll details
			strSQL = "SELECT " & strDbTable & "EmailNotify.*  " & _
			"FROM " & strDbTable & "EmailNotify" & strRowLock & " " & _
			"WHERE " & strDbTable & "EmailNotify.Author_ID = " & lngLoggedInUserID & " AND " & strDbTable & "EmailNotify.Forum_ID=" & intForumID & ";"
		
		
			'Set the cursor type property of the record set to Forward Only
			.CursorType = 0
		
			'Set the Lock Type for the records so that the record set is only locked when it is updated
			.LockType = 3
				
			'Query the database
			.Open strSQL, adoCon
		
			'If the user no-longer wants email notification for this topic then remove the entry form the db
			If NOT .EOF Then
		
				Do while NOT .EOF
					'Delete the db entry
					.Delete
					
					'Move to next record
					.MoveNext
						
					'Set the return value
					strReturnValue = "&EN=FU"
					strReturnPage = "forum_topics.asp?FID=" & intForumID & strReturnValue
				Loop
			End If
		
		End If

		'Clean up
		.Close

	End With
	
	'Clean up
	Call closeDatabase()
	
	'If there is no return page value then return to the forum
	If strReturnPage = "" Then
		Response.Redirect("forum_topics.asp?FID=" & intForumID & strQsSID3)
	'Else return just to forum or topic that is related to
	Else
		Response.Redirect(strReturnPage & strQsSID3)
	End If
End Function
%>