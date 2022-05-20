<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="admin_common.asp" -->
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



'Set the timeout of the page
Server.ScriptTimeout =  1000


'Set the response buffer to true as we maybe redirecting
Response.Buffer = True




'Dimension variables
Dim rsTopic		'Holds the database recordset for the Topic to be delete
Dim intCatID		'Holds the forum ID to be deleted
Dim lngPollID		'Holds the poll ID if there is one to delete
Dim rsPost		'Recordset to get the post ID from db to check against guest post names
Dim saryFileUploads	'Holds the files to be deleted
Dim intLoop		'Loop counter
Dim objFSO		'Holds the FSO object




'Check the session key ID to prevent XCSRF
Call checkFormID(Request.QueryString("XID"))


'Get the category ID to delete
intCatID = IntC(Request.QueryString("CatID"))



'Get all the Forums from the database to be deleted

'Create a recordset object for the Topic in the database
Set rsTopic = Server.CreateObject("ADODB.Recordset")
Set rsPost = Server.CreateObject("ADODB.Recordset")


'Initalise the strSQL variable with an SQL statement to get the topic from the database
strSQL = "SELECT " & strDbTable & "Forum.* " & _
"FROM " & strDbTable & "Forum " & _
"WHERE " & strDbTable & "Forum.Cat_ID="  & intCatID & ";"

'Query the database
rsCommon.Open strSQL, adoCon



'Loop through all the forums for the topics and delete them
Do While NOT rsCommon.EOF

	'Read in the forum ID
	intForumID = CInt(rsCommon("Forum_ID"))

	'Get all the Topics from the database to be deleted
	'Initalise the strSQL variable with an SQL statement to get the topic from the database
	strSQL = "SELECT " & strDbTable & "Topic.Topic_ID, " & strDbTable & "Topic.Poll_ID " & _
	"FROM " & strDbTable & "Topic " & _
	"WHERE " & strDbTable & "Topic.Forum_ID ="  & intForumID & ";"

	'Query the database
	rsTopic.Open strSQL, adoCon

	'Loop through all the threads for the topics and delete them
	Do While NOT rsTopic.EOF
	
	
	
		'First we need to delete any entry in the GuestName table incase this was a guest poster posting the message
		
		'Initalise the strSQL variable with an SQL statement to get the topic from the database
		strSQL = "SELECT " & strDbTable & "Thread.Thread_ID " & _
		"FROM " & strDbTable & "Thread " & _
		"WHERE " & strDbTable & "Thread.Topic_ID=" & rsTopic("Topic_ID") & ";"
		
		'Query the database
		rsPost.Open strSQL, adoCon
		
		'Loop through thread ID's
		Do While NOT rsPost.EOF
		
			'First we need to delete any entry in the GuestName table incase this was a guest poster posting the message
			strSQL = "DELETE FROM " & strDbTable & "GuestName WHERE " & strDbTable & "GuestName.Thread_ID=" & CLng(rsPost("Thread_ID")) & ";"
		
			'Excute SQL
			adoCon.Execute(strSQL)
			
			'Movenext rs
			rsPost.MoveNext
		Loop
		
		'Close rs
		rsPost.Close
	
	

		'Delete the posts in this topic
		strSQL = "DELETE FROM " & strDbTable & "Thread WHERE " & strDbTable & "Thread.Topic_ID ="  & CLng(rsTopic("Topic_ID")) & ";"

		'Write to database
		adoCon.Execute(strSQL)


		'Delete any polls

		'Get the Poll ID
		lngPollID = CLng(rsTopic("Poll_ID"))

		'If there is a poll delete that as well
		If lngPollID > 0 Then

			'Delete the poll choice
			strSQL = "DELETE FROM " & strDbTable & "PollChoice WHERE " & strDbTable & "PollChoice.Poll_ID=" & lngPollID & ";"

			'Delete the threads
			adoCon.Execute(strSQL)
			
			'Delete the Poll Votes 
			strSQL = "DELETE FROM " & strDbTable & "PollVote " & strRowLock & " WHERE " & strDbTable & "PollVote.Poll_ID=" & lngPollID & ";" 
				
			'Write to database 
			adoCon.Execute(strSQL)

			'Delete the poll choice
			strSQL = "DELETE FROM " & strDbTable & "Poll WHERE " & strDbTable & "Poll.Poll_ID=" & lngPollID & ";"

			'Delete the threads
			adoCon.Execute(strSQL)
		End If


		'Move to the next record
		rsTopic.MoveNext
	Loop




	'Delete the topics in this forum
	strSQL = "DELETE FROM " & strDbTable & "Topic WHERE " & strDbTable & "Topic.Forum_ID = "  & intForumID & ";"

	adoCon.Execute(strSQL)


	'Delete any group permissions set for the forum
	strSQL = "DELETE FROM " & strDbTable & "Permissions WHERE " & strDbTable & "Permissions.Forum_ID = "  & intForumID & ";"
	
	adoCon.Execute(strSQL)
	
	
	'Delete email notifications
	strSQL = "DELETE FROM " & strDbTable & "EmailNotify WHERE " & strDbTable & "EmailNotify.Forum_ID = "  & intForumID & ";"
	
	adoCon.Execute(strSQL)


	'close recordset
	rsTopic.Close



	'Move to the next record
	rsCommon.MoveNext
Loop



'Delete the forum
strSQL = "DELETE FROM " & strDbTable & "Forum WHERE " & strDbTable & "Forum.Cat_ID = "  & intCatID & ";"

'Write to database
adoCon.Execute(strSQL)



'Delete the category
strSQL = "DELETE FROM " & strDbTable & "Category WHERE " & strDbTable & "Category.Cat_ID = "  & intCatID & ";"

'Write to database
adoCon.Execute(strSQL)



'Reset Server Objects
Set rsPost = Nothing
rsCommon.Close
Set rsTopic = Nothing
Call closeDatabase()


'Return to the forum categories page
Response.Redirect("admin_view_forums.asp" & strQsSID1)
%>