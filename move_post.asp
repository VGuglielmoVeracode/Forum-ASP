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
Dim strNewTopicSubject	'Holds the new subject
Dim lngOldTopicID	'Holds the old topic ID number
Dim lngPostID		'Holds the post ID
Dim strPostDateTime	'Holds the date of the post


'If the user is user is using a banned IP redirect to an error page
If bannedIP() Then
	
	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp?M=IP" & strQsSID3)
End If


'Check the form ID to prevent XCSRF
Call checkFormID(Request.Form("formID"))



'Read in the post ID
lngPostID = LngC(Request.Form("PID"))


'Query the datbase to get the forum ID for this post
strSQL = "SELECT " & strDbTable & "Topic.Forum_ID, " & strDbTable & "Topic.Topic_ID, " & strDbTable & "Thread.Message_date " & _
"FROM " & strDbTable & "Topic" & strDBNoLock & ", " & strDbTable & "Thread" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Topic.Topic_ID = " & strDbTable & "Thread.Topic_ID " & _
	"AND " & strDbTable & "Thread.Thread_ID = " & lngPostID & ";"

'Query the database
rsCommon.Open strSQL, adoCon


'If there is a record returened read in the forum ID
If NOT rsCommon.EOF Then
	intForumID = CInt(rsCommon("Forum_ID"))
	lngOldTopicID = CLng(rsCommon("Topic_ID"))
	strPostDateTime = CDate(rsCommon("Message_date"))
End If

'Clean up
rsCommon.Close


'Call the moderator function and see if the user is a moderator
If blnAdmin = False Then blnModerator = isModerator(intForumID, intGroupID)


'If the user is not a moderator or admin then keck em
If blnAdmin = false AND  blnModerator = false Then
	
	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)


'The person is an admin of modertor so move the post
Else

	'Read in the forum details
	lngTopicID = LngC(Request.Form("topicSelect"))
	intForumID = IntC(Request.Form("toFID"))
	strNewTopicSubject = Request.Form("subject")




	'If a new subject has been entered then place it into the database
	If strNewTopicSubject <> "" Then

		'Get rid of scripting tags in the subject
		strNewTopicSubject = removeAllTags(strNewTopicSubject)
		strNewTopicSubject = formatInput(strNewTopicSubject)

		'Initalise the SQL string with a query to get the Topic details
		strSQL = "SELECT" & strDBTop1 & " " & strDbTable & "Topic.Forum_ID, " & strDbTable & "Topic.Topic_ID, " & strDbTable & "Topic.Subject, " & strDbTable & "Topic.Start_Thread_ID, " & strDbTable & "Topic.Last_Thread_ID, " & strDbTable & "Topic.Poll_ID, " & strDbTable & "Topic.Priority, " & strDbTable & "Topic.Hide " & _
		"FROM " & strDbTable & "Topic" & strRowLock & " " & _
		"WHERE Forum_ID = " & intForumID & " " & _
		"ORDER BY " & strDbTable & "Topic.Topic_ID DESC" & strDBLimit1 & ";"

		With rsCommon
			'Set the cursor type property of the record set to Forward Only
			.CursorType = 0

			'Set the Lock Type for the records so that the record set is only locked when it is updated
			.LockType = 3

			'Open the topic table
			.Open strSQL, adoCon

			'Insert the new topic details in the recordset
			.AddNew

			.Fields("Forum_ID") = intForumID
			.Fields("Subject") = strNewTopicSubject
			.Fields("Start_Thread_ID") = -1 '-1 is set at the moment as we need a value and to prevent erros with moving hidden posts, the topic stats function is run at the end to correct this
			.Fields("Last_Thread_ID") = -1 'same as above
			.Fields("Poll_ID") = 0
			.Fields("Priority") = 0
			.Fields("Hide") = False
			

			'Update the database with the new topic details
			.Update

			'Re-run the Query once the database has been updated
			.Requery

			'Read in the new topic's ID number
			lngTopicID = CLng(rsCommon("Topic_ID"))

			'Clean up
			.Close
		End With
	End If



		
	'Move the post to another topic use ADO with requery otherwise access can be to slow on some servers coursing issues with incorrect stats for topics
	If strDatabaseType = "Access" Then
		strSQL = "SELECT " & strDbTable & "Thread.Topic_ID " & _
		"FROM " & strDbTable & "Thread" & strRowLock & " " & _
		"WHERE " & strDbTable & "Thread.Thread_ID = " & lngPostID & ";"
	
		With rsCommon
		
			'Set the cursor type property of the record set to Forward Only
			.CursorType = 0
			
			'Set the Lock Type for the records so that the record set is only locked when it is updated
			.LockType = 3
			
			'Open the thread table
			.Open strSQL, adoCon
			
			'If the record is returned update the topic ID
			If NOT .EOF Then
				
				.Fields("Topic_ID") = lngTopicID
			End If
			
			'Update the database
			.Update
			
			'Requery the db for slow old access
			.Requery
			
			'Clean up
			.Close
		End With
	
	
	'Else use the faster SQL update to move the post, also mySQL dosn't like the ADO method used for Access
	Else
		'Initliase SQL upadte
		strSQL = "UPDATE " & strDbTable & "Thread " & _
		"SET " & strDbTable & "Thread.Topic_ID = " & lngTopicID & " " & _
		"WHERE " & strDbTable & "Thread.Thread_ID = "  & lngPostID & ";"
		
		'Execute SQL
		adoCon.Execute(strSQL)	
	End If
	
	
	
	
	'Check there are still posts in the old topic, if not delete the old topic
	With rsCommon
		
		strSQL = "SELECT" & strDBTop1 & " " & strDbTable & "Thread.Thread_ID " & _
		"FROM " & strDbTable & "Thread" & strDBNoLock & " " & _
		"WHERE Topic_ID = " & lngOldTopicID & strDBLimit1 & ";"
	
		'Open the thread table
		.Open strSQL, adoCon
	
		'See if there is a topic left in the old topic
		If .EOF Then
			'If there are no topics left then delete the old topic
			strSQL = "DELETE FROM " & strDbTable & "Topic WHERE Topic_ID = " & lngOldTopicID & ";"
	
			'Write to database
			adoCon.Execute(strSQL)
		End If
	
		'Close the recordset
		.Close
	End With
	
	
	'Update the forum stats to get the topics in the correct order etc.
	Call updateTopicStats(lngOldTopicID)
	Call updateTopicStats(lngTopicID)
	
	'Call again as some systems are slow
	Call updateTopicStats(lngOldTopicID)
	Call updateTopicStats(lngTopicID)

End If

'Reset main server variables
Call closeDatabase()
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<script language="JavaScript">
	window.opener.location.href = 'forum_posts.asp?TID=<% = lngTopicID %><% = strQsSID2 %>'
	window.close();
</script>
</head>
</html>