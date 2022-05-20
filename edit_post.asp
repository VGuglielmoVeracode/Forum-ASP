<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="functions/functions_send_mail.asp" -->
<!--#include file="functions/functions_format_post.asp"	-->
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




'Set the response buffer to true as we maybe redirecting
Response.Buffer	= True


'If in demo mode redirect
If blnDemoMode Then
	Call closeDatabase()
	Response.Redirect("insufficient_permission.asp?M=DEMO" & strQsSID3)
End If


'Dimension variables
Dim blnEmailNotify		'Set to	true if	the users want to be notified by e-mail	of a post
Dim strMessage			'Holds the Users Message
Dim lngMessageID		'Holds the message ID number
Dim strMode			'Holds the mode	of the page so we know whether we are editing, updating, or new	topic
Dim lngTopicID			'Holds the topic ID number
Dim strSubject			'Holds the subject
Dim blnSignature		'Holds wether a	signature is to	be shown or not
Dim intPriority			'Holds the priority of tipics
Dim intReturnPageNum		'Holds the page	number to return to
Dim strReturnCode		'Holds the code	if the post is not valid and we	need to	return to forum	without	posting
Dim strPollQuestion		'Holds the poll	question
Dim blnMultipleVotes		'Set to	true if	multiple votes are allowed
Dim blnPollReply		'Set to	true if	users can't reply to a poll
Dim saryPollChoice()		'Array to hold the poll	choices
Dim intPollChoice		'Holds the poll	choices	loop counter
Dim strBadWord			'Holds the bad words
Dim strBadWordReplace		'Holds the rplacment word for the bad word
Dim lngPollID			'Holds the poll	ID number
Dim blnForumLocked		'Set to true if the forum is locked
Dim blnTopicLocked		'Set to true if the topic is locked
Dim strGuestName		'Holds the name of the guest if it is a guest posting
Dim lngStartThreadID		'Holds the thread ID of the first post in the topic to use for security checking
Dim saryFileUploads		'Holds the names of the files uploaded
Dim objFSO			'Holds the file system object
Dim intLoop			'Loop counter
Dim strTopicIcon		'Holds the topic icon for the message
Dim dtmEventDate		'Holds the Calendar event date
Dim dtmEventDateEnd		'Holds the Calendar event date
Dim dtmMessageDateTime		'Holds the date and time of the post
Dim lngAuthorID			'Holds the author ID of the person who created the post
Dim objRegExp			'used for searches


'Initalise variables
lngPollID = 0
blnForumLocked = False
blnTopicLocked = False
blnCheckFirst = False

'If the	user has not logged in then redirect them to the main forum page
If lngLoggedInUserID = 0 OR blnActiveMember = False OR blnBanned Then
	
	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("default.asp" & strQsSID1)
End If



'******************************************
'***	      Check IP address		***
'******************************************

'If the	user is	user is	using a	banned IP redirect to an error page
If bannedIP() Then
	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp?M=IP" & strQsSID3)

End If


'Check the session ID to stop spammers using the email form
Call checkFormID(Request.Form("formID"))




'******************************************
'***	  Read in form details		***
'******************************************

'Read in user deatils from the post message form
strMode	= Trim(Mid(Request.Form("mode"), 1, 10))
intForumID = IntC(Request.Form("FID"))
lngTopicID = LngC(Request.Form("TID"))
strSubject = Trim(Mid(Request.Form("subject"), 1, 50))
strMessage = Mid(Request.Form("Message"), 1, 200000)
lngMessageID = LngC(Request.Form("PID"))
blnEmailNotify = BoolC(Request.Form("email"))
blnSignature = BoolC(Request.Form("signature"))
intPriority = IntC(Request.Form("priority"))
strTopicIcon =  Request.Form("icon")
'If the user is in a guest then get there name
If lngLoggedInUserID = 2 Then strGuestName = Trim(Mid(Request.Form("Gname"), 1, 20))

'Read in Calendar event date
If Request.Form("eventDay") <> 0 AND Request.Form("eventMonth") <> 0 AND Request.Form("eventYear") <> 0 Then
	dtmEventDate = internationalDateTime(DateSerial(Request.Form("eventYear"), Request.Form("eventMonth"), Request.Form("eventDay")))
End If

'Read in event end date
If Request.Form("eventDayEnd") <> 0 AND Request.Form("eventMonthEnd") <> 0 AND Request.Form("eventYearEnd") <> 0 Then
	dtmEventDateEnd = internationalDateTime(DateSerial(Request.Form("eventYearEnd"), Request.Form("eventMonthEnd"), Request.Form("eventDayEnd")))

	'If the end date is before the start date don't add it to the database
	If dtmEventDate => dtmEventDateEnd OR dtmEventDate = "" Then dtmEventDateEnd = null
End If



'******************************************
'***	     Get permissions	      *****
'******************************************


'Get the forum permissions from the topic being posted in and also check if the topic is locked and who posted the topic
strSQL = " " & _
"SELECT" & strDBTop1 & " " & strDbTable & "Forum.Password, " & strDbTable & "Forum.Forum_code, " & strDbTable & "Forum.Locked AS ForumLocked, " & strDbTable & "Forum.Password, " & strDbTable & "Topic.Locked AS TopicLocked, " & strDbTable & "Topic.Poll_ID, " & strDbTable & "Topic.Start_Thread_ID, " & strDbTable & "Permissions.* " & _
"FROM " & strDbTable & "Forum" & strDBNoLock & ", " & strDbTable & "Topic" & strDBNoLock & ", " & strDbTable & "Permissions" & strDBNoLock & " " & _
"WHERE  " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Topic.Forum_ID " & _
	"AND " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Permissions.Forum_ID " & _
	"AND " & strDbTable & "Topic.Topic_ID = " & lngTopicID & " " & _
	"AND (" & strDbTable & "Permissions.Author_ID = " & lngLoggedInUserID & " OR " & strDbTable & "Permissions.Group_ID = " & intGroupID & ") " & _
"ORDER BY " & strDbTable & "Permissions.Author_ID DESC" & strDBLimit1 & ";"

'Query the database
rsCommon.Open strSQL, adoCon



'Check the forum permissions
If NOT rsCommon.EOF Then

	'Get forum ID
	intForumID = CInt(rsCommon("Forum_ID"))
	
	'If this isn't the first post in the topic then it is just a plain edit and NOT a poll or topic subject edit!!
	If lngMessageID <> CLng(rsCommon("Start_Thread_ID")) Then strMode = "edit"
	
	'Get the POLL ID if there is a poll to be edited
	If strMode = "editPoll" Then lngPollID = CLng(rsCommon("Poll_ID"))
	
	'See if the topic is locked if this is not the admin
	If blnAdmin = False Then blnTopicLocked = CBool(rsCommon("TopicLocked"))
	
	'See if the forum is locked if this is not the admin
	If blnAdmin = False Then blnForumLocked = CBool(rsCommon("ForumLocked"))

	'Read in the forum permissions
	blnRead = CBool(rsCommon("View_Forum"))
	blnEdit = CBool(rsCommon("Edit_posts"))
	blnPriority = CBool(rsCommon("Priority_posts"))
	blnPollCreate = CBool(rsCommon("Poll_create"))
	blnModerator = CBool(rsCommon("Moderate"))
	blnEvents = CBool(rsCommon("Calendar_event"))
	
	
	'If this is a modertor then make sure they have edit rights
	If blnAdmin OR blnModerator Then blnEdit = true
		
	'If this in not an admin or moderator set the priority to 0
	If (blnAdmin = false OR blnModerator = false) AND blnPriority = false Then intPriority = 0
		
		
	
	'If the user has no read or edit rights then kick them
	If blnRead = False OR blnEdit = False Then
		
		'Reset Server Objects
		rsCommon.Close
		Call closeDatabase()


		'Redirect to a page asking for the user to enter the forum password
		Response.Redirect("insufficient_permission.asp" & strQsSID1)
	End If


	'If the forum requires a password and a logged in forum code is not found on the users machine then send them to a login page
	If rsCommon("Password") <> "" AND (getCookie("fID", "Forum" & intForumID) <> rsCommon("Forum_code") AND getSessionItem("FP" & intForumID) <> "1") Then

		'Reset Server Objects
		rsCommon.Close
		Call closeDatabase()

		'Redirect to a page asking for the user to enter the forum password
		Response.Redirect("forum_password_form.asp?FID=" & intForumID & strQsSID3)
	End If
	
	
	'If this is the admin or moderator then set the post to be displayed
	If blnAdmin OR blnModerator Then blnCheckFirst = false

End If

'Clean up
rsCommon.Close






'******************************************
'***	     SPAM Filters		***
'******************************************

'Get the spam filters
'Initalise the SQL string with a query to read in all the words from the smut table
strSQL = "SELECT " & strDbTable & "Spam.* FROM " & strDbTable & "Spam" & strDBNoLock & ";"
	
'Open the recordset
rsCommon.Open strSQL, adoCon
		
'Create regular experssions object
Set objRegExp = New RegExp

'Loop through all the spam filters
Do While NOT rsCommon.EOF
			
	'Tell the regular experssions object what to look for
	With objRegExp
		.Pattern = rsCommon("Spam")
		.IgnoreCase = True
		.Global = True
	End With
		
	'Ignore errors, incase someone entered an incorrect bad word that breakes regular expressions
	On Error Resume Next

	'See if the spam is in the message
	If objRegExp.Execute(strMessage).Count > 0 Then 
		
		'If wqe need to reject kick the user
		If rsCommon("Spam_Action") = "Reject" Then
			strReturnCode =	"spam"
		
		'Else set the message to require approval
		Else
			blnCheckFirst = True
		End If	
		
	End If
				
			
	'Disable error trapping
	On Error goto 0
	
	'Move to the next word in the recordset
	rsCommon.MoveNext
Loop
		
'Distroy regular experssions object
Set objRegExp = nothing
	
'Release server objects
rsCommon.Close








'*****************************************************
'***   Redirect if the forum or topic is locked   ****
'*****************************************************

'If the forum or topic is locked then don't let the user post a message
If blnForumLocked OR blnTopicLocked Then

	'Clean up
	Call closeDatabase()

	'Redirect to error page
	If blnForumLocked Then
		Response.Redirect("not_posted.asp?mode=FLocked" & strQsSID3)
	Else
		Response.Redirect("not_posted.asp?mode=TClosed" & strQsSID3)
	End If
End If








'******************************************
'***	 Get return page details      *****
'******************************************

'If there is no	number must be a new post
If NOT isNumeric(Request.Form("PN")) Then
	intReturnPageNum = 1
Else
	intReturnPageNum = IntC(Request.Form("PN"))
End If

'calcultae which page the tread	is posted on
If Request.Form("ThreadPos") <> "" Then

	'If the	position in the	topic is on next page add 1 to the return page number
	If IntC(Request.Form("ThreadPos")) > (intThreadsPerPage	* intReturnPageNum) Then
		intReturnPageNum = intReturnPageNum + 1
	End If
End If



'********************************************
'***  Clean up and check in form details  ***
'********************************************

'Remove any tabs from subject
strSubject = Replace(strSubject, vbTab, " ")

'Fix for UK Pound Sterling displaying as Â£
strSubject = Replace(strSubject, Chr(194) & Chr(163), "£", 1, -1, 1)

'If there is no	subject	or message then	don't post the message as won't	be able	to link	to it
If strSubject =	"" AND (strMode = "editTopic" OR strMode = "poll") Then strReturnCode = "noSubject"
If Trim(strMessage) = "" OR Trim(strMessage) = "<P>&nbsp;</P>" OR Trim(strMessage) = "<br>" OR Trim(strMessage) = "<br>" & vbCrLf Then strReturnCode = "noSubject"



'Place format posts posted with	the WYSIWYG Editor (RTE)
If Request.Form("browser") = "RTE" Then

	'Call the function to format WYSIWYG posts
	strMessage = WYSIWYGFormatPost(strMessage)

'Else standrd editor is	used so	convert	forum codes
Else
	'Call the function to format posts
	strMessage = FormatPost(strMessage)
End If


'If the user wants forum codes enabled then format the post using them
If Request.Form("forumCodes") Then strMessage = FormatForumCodes(strMessage)
	

'Convert text links to hyperlinks
If blnCovertTextLinkstoHyperLink AND blnGroupURLs Then
		
	'If there are no hyperlinks already look to convert text links to hyperlinks
	If InStr(1, strMessage, "<a", 1) = 0 AND InStr(1, strMessage, "<img", 1) = 0 Then
		strMessage = InsertHyperlinks(strMessage)
	End If
		
End If



'Check the message for malicious HTML code
strMessage = HTMLsafe(strMessage)




'If the user is in a guest then clean up their username to remove malicious code
If lngLoggedInUserID = 2 Then
	strGuestName = formatSQLInput(strGuestName)
	strGuestName = formatInput(strGuestName)
End If



'If topic icons then clean up any input
If blnTopicIcon Then
	
	'If the topic icon is not selected don't fill the db with crap and leave field empty
	If strTopicIcon = strImagePath & "blank_smiley.gif" Then strTopicIcon = ""
	
	'Clean up user input
	strTopicIcon = formatInput(strTopicIcon)
	strTopicIcon = removeAllTags(strTopicIcon)
End If





'********************************************
'***	Read in	poll details (if Poll)	  ***
'********************************************

'If this is a poll then read in the poll details
If strMode = "editPoll" AND lngPollID > 0 Then

	'Read in poll question and multiple votes
	strPollQuestion	= Trim(Mid(Request.Form("pollQuestion"), 1, 70))
	blnMultipleVotes = BoolC(Request.Form("multiVote"))
	blnPollReply = BoolC(Request.Form("pollReply"))

	'If there is no	poll question then there initilise the error variable
	If strPollQuestion = ""	Then strReturnCode = "noPoll"

	'Clean up poll question
	strPollQuestion	= removeAllTags(strPollQuestion)


	'Loop through and read in the poll question
	For intPollChoice = 1 To intMaxPollChoices

		'ReDimension the array for the correct number of choices
		'ReDimensioning	arrays is bad for performance but usful	in this	for what I need	it for
		ReDim Preserve saryPollChoice(intPollChoice)

		'Read in the poll choice
		saryPollChoice(intPollChoice) =	Trim(Mid(Request.Form("choice" & intPollChoice), 1, 60))

		'If there is nothing in	position 1 and 2 set a return error code
		If intPollChoice < 2 AND saryPollChoice(intPollChoice) = "" Then strReturnCode = "noPoll"


		'Clean up input
		saryPollChoice(intPollChoice) =	removeAllTags(saryPollChoice(intPollChoice))
	Next
End If





'******************************************
'***	     Filter Bad	Words	      *****
'******************************************

'Initalise the SQL string with a query to read in all the words	from the smut table
strSQL = "SELECT " & strDbTable & "Smut.* " & _
"FROM " & strDbTable & "Smut " & strDBNoLock & ";"

'Open the recordset
rsCommon.Open strSQL, adoCon

'Create regular experssions object
Set objRegExp = New RegExp

'Loop through all the words to check for
Do While NOT rsCommon.EOF

	'Put the bad word into a string	for imporoved perfoamnce
	strBadWord = rsCommon("Smut")
	strBadWordReplace = rsCommon("Word_replace")

	'Tell the regular experssions object what to look for
	With objRegExp
		.Pattern = strBadWord
		.IgnoreCase = True
		.Global = True
	End With
	
	'Ignore errors, incase someone entered an incorrect bad word that breakes regular expressions
	On Error Resume Next
	
	'Replace the swear words with the words	in the database	the swear words
	strSubject = objRegExp.Replace(strSubject, strBadWordReplace)
	strMessage = objRegExp.Replace(strMessage, strBadWordReplace)
	
	'Disable error trapping
	On Error goto 0

	'If this is a poll run the poll	choices	through	the bad	word filter as well
	If strMode = "poll" Then

		'Clean up the poll question
		strPollQuestion	= objRegExp.Replace(strPollQuestion, strBadWordReplace)

		'Loop though and check all the strings in the Poll array
		For intPollChoice = 1 To UBound(saryPollChoice)
		
			'Ignore errors, incase someone entered an incorrect bad word that breakes regular expressions
			On Error Resume Next
	
			saryPollChoice(intPollChoice) =	objRegExp.Replace(saryPollChoice(intPollChoice), strBadWordReplace)
			
			'Disable error trapping
			On Error goto 0
		Next
	End If
	
	

	'Move to the next word in the recordset
	rsCommon.MoveNext
Loop

'Distroy regular experssions object
Set objRegExp = nothing

'Reset server varaible
rsCommon.Close







'Get rid of scripting tags in the subject
'This is done after the bad word filter incase the forum admin is replacing bad words with HTML content
strSubject = removeAllTags(strSubject)





'**********************************************
'***  If input problems	send to	error page  ***
'**********************************************

'If there is a return code then	this post is not valid so redirect to error page
If strReturnCode <> "" Then

	'Clean up
	Call closeDatabase()

	'Redirect to error page
	Response.Redirect("not_posted.asp?mode=" & strReturnCode &  strQsSID3)
End If







'******************************************
'***	    Edit Post Update		***
'******************************************



'Initalise the strSQL variable with an SQL statement to	query the database get the message details
strSQL = "SELECT " & strDbTable & "Thread.Thread_ID, " & strDbTable & "Thread.Author_ID, " & strDbTable & "Thread.Message, " & strDbTable & "Thread.Show_signature, " & strDbTable & "Thread.IP_addr, " & strDbTable & "Thread.Hide, " & strDbTable & "Thread.Message_date " & _
"FROM	" & strDbTable & "Thread" & strRowLock & " " & _
"WHERE " & strDbTable & "Thread.Thread_ID = " & lngMessageID & ";"


'Set the cursor	type property of the record set	to Forward only
rsCommon.CursorType = 0

'Set the Lock Type for the records so that the record set is only locked when it is updated
rsCommon.LockType = 3

'Open the author table
rsCommon.Open strSQL, adoCon


'Read in the author ID
lngAuthorID = CLng(rsCommon("Author_ID"))


'Only update the post if this is a moderator, forum admin, or the person who posted
If (blnAdmin OR blnModerator) OR (lngAuthorID = lngLoggedInUserID) Then
	
	'Read in the message date
	dtmMessageDateTime = CDate(rsCommon("Message_date"))
	
	
	'See if editing is within the time limit
	If intEditPostTimeFrame > 0 AND (blnAdmin = False AND blnModerator = False) Then
		If DateDiff("n", dtmMessageDateTime, Now()) >= intEditPostTimeFrame Then 
			
			'Reset Server Objects
			rsCommon.Close
			Call closeDatabase()
		
			'Redirect to a page asking for the user to enter the forum password
			Response.Redirect("insufficient_permission.asp?M=eExp" & strQsSID3)
			
		End If
	End If
	
	
	'If we are to show who edit the post and time then contantinet it to the end of the message
	'Update in 10 to only use this if the post is over xx minutes old
	If blnShowEditUser AND DateDiff("n", dtmMessageDateTime, Now()) >= intEditedTimeDelay Then
		strMessage = strMessage & "<edited><editID>" & strLoggedInUsername & "</editID><editDate>" & internationalDateTime(Now()) &  "</editDate></edited>"
	End If
	

	'If this is a normal user let 'em know their post needs to be checked first before it is displayed (if hidden)
	If (blnAdmin = false OR blnModerator = false) AND blnCheckFirst = False Then blnCheckFirst = CBool(rsCommon("Hide"))

	'Enter the updated post	into the recordset
	rsCommon.Fields("Message") = strMessage
	rsCommon.Fields("Show_signature") = CBool(blnSignature)
	'Only update the IP address if this is not the admin
	If blnAdmin = False AND blnModerator = False  Then rsCommon.Fields("IP_addr") = getIP()
	rsCommon.Fields("Hide") = blnCheckFirst
	

	'Update	the database
	rsCommon.Update

	'Close rs
	rsCommon.Close

'Else the user does not have permission to edit this post/topic/poll, so kick 'em
Else

	'Reset Server Objects
	rsCommon.Close
	Call closeDatabase()


	'Redirect to a page asking for the user to enter the forum password
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If









'********************************************
'***		  Edit Poll		  ***
'********************************************

'If this is a poll then save the poll to the database
If strMode = "editPoll" AND lngPollID > 0 Then

	'********************************************
	'***	     Update poll question		  ***
	'********************************************

	'Initalise the SQL string with a query to get the poll last poll details to get the poll ID number in next (use nolock as this is a new insert so a dirty read is OK)
	strSQL = "SELECT " & strDbTable & "Poll.* " & _
	"FROM " & strDbTable & "Poll" & strRowLock & " " & _
	"WHERE " & strDbTable & "Poll.Poll_ID=" & lngPollID & ";"

	With rsCommon
		'Set the cursor	type property of the record set	to Forward only
		.CursorType = 0

		'Set the Lock Type for the records so that the record set is only locked when it is updated
		.LockType = 3

		'Open the poll table
		.Open strSQL, adoCon

		'Update	recordset
		.Fields("Poll_question") = strPollQuestion
		.Fields("Multiple_votes") = blnMultipleVotes
		.Fields("Reply") = blnPollReply

		'Update	the database with the new poll question
		.Update

		'Clean up
		.Close
	End With


	'********************************************
	'***	      Update poll choices	  ***
	'********************************************

	'Initalise the SQL string with a query to get the choice 
	strSQL = "SELECT " & strDbTable & "PollChoice.Poll_ID, " & strDbTable & "PollChoice.Choice " & _
	"FROM " & strDbTable & "PollChoice" & strRowLock & " " & _
	"WHERE " & strDbTable & "PollChoice.Poll_ID=" & lngPollID & ";"

	With rsCommon
		'Set the cursor	type property of the record set	to Forward only
		.CursorType = 0

		'Set the Lock Type for the records so that the record set is only locked when it is updated
		.LockType = 3

		'Open the author table
		.Open strSQL, adoCon
		
		intPollChoice = 0

		'Add the new poll choices to recordset
		Do While NOT .EOF
			
			'Move to next poll choice
			If intPollChoice < UBound(saryPollChoice) Then intPollChoice = intPollChoice + 1
		
			'Update	recordset
			.Fields("Choice") = saryPollChoice(intPollChoice)
			
			'Update	the database with the poll choices (bad place to do it but this prevents errors)
			.Update
			
			'Move to next record
			.MoveNext
		Loop

		'Clean up
		.Close
	End With

	'Change	the mode to editTopic to save any updated topic subject
	strMode = "editTopic"
End If






'******************************************
'***	     Edit Topic	Update		***
'******************************************

'If the	post is	the first in the thread	then update the	topic details
If strMode = "editTopic" Then

	'Initalise the SQL string with a query to get the Topic	details
	strSQL = "SELECT " & strDbTable & "Topic.Subject, " & strDbTable & "Topic.Icon, " & strDbTable & "Topic.Priority, " & strDbTable & "Topic.Event_date, " & strDbTable & "Topic.Event_date_end " & _
	"FROM " & strDbTable & "Topic" & strRowLock & " " & _
	"WHERE " & strDbTable & "Topic.Topic_ID=" & lngTopicID & ";"

	With rsCommon
		'Set the cursor	type property of the record set	to Forward only
		.CursorType = 0

		'Set the Lock Type for the records so that the record set is only locked when it is updated
		.LockType = 3

		'Open the author table
		.Open strSQL, adoCon

		'Update	the recorset
		.Fields("Subject") = strSubject
		If blnTopicIcon Then .Fields("Icon") = strTopicIcon
		.Fields("Priority") = intPriority
		'If Calendar events allowed save
		If blnCalendar AND blnEvents Then .Fields("Event_date") = dtmEventDate
		If blnCalendar AND blnEvents Then .Fields("Event_date_end") = dtmEventDateEnd

		'Update	the database with the new topic	details
		.Update

		'Clean up
		.Close
	End With
End If





'******************************************
'***	     Logging		***
'******************************************

'If logging enabled log who edited topic/post
If blnLoggingEnabled AND (blnEditPostLogging OR (blnModeratorLogging AND (blnAdmin OR blnModerator))) Then
	
	'Initalise the SQL string with a query to get the Topic	details
	strSQL = "SELECT " & strDbTable & "Topic.Subject " & _
	"FROM " & strDbTable & "Topic " & _
	"WHERE " & strDbTable & "Topic.Topic_ID = " & lngTopicID & ";"
	
	'Open the author table
	rsCommon.Open strSQL, adoCon

	Call logAction(strLoggedInUsername, "Edited Post in '" & decodeString(rsCommon("Subject")) & "' - PostID " & lngMessageID)

	
	rsCommon.Close
End If




'**********************************************************
'***	     Update Email Notify if this is a reply	***
'**********************************************************

'Delete	or Save	email notification for the user, if email notify is enabled

If blnEmail = True Then

	'Initalise the SQL string with a query to get the email	notify details
	strSQL = "SELECT " & strDbTable & "EmailNotify.*  " & _
	"FROM	" & strDbTable & "EmailNotify" & strRowLock & " " & _
	"WHERE " & strDbTable & "EmailNotify.Author_ID=" & lngLoggedInUserID & " "  & _
		"AND " & strDbTable & "EmailNotify.Topic_ID=" & lngTopicID & ";"


	With rsCommon

		'Set the cursor	type property of the record set	to Forward only
		.CursorType = 0

		'Set the Lock Type for the records so that the record set is only locked when it is updated
		.LockType = 3

		'Query the database
		.Open strSQL, adoCon


		'If the	user no-longer wants email notification	for this topic then remove the entry form the db
		If blnEmailNotify = False AND NOT .EOF Then

			'Delete	the db entry
			.Delete

		'Else if this is a new post and	the user wants to be notified add the new entry	to the database
		ElseIf blnEmailNotify = True AND .EOF Then

			'Add new rs
			.AddNew

			'Create	new entry
			.Fields("Author_ID") = lngLoggedInUserID
			.Fields("Topic_ID") = lngTopicID

			'Upade db with new rs
			.Update
		End If

		'Clean up
		.Close

	End With
End If




'******************************************
'***	    Clean up objects		***
'******************************************

'Reset Server Objects
Call closeDatabase()


'If the sort order has been changed for this sesison then update the Page Number (PN)
If getSessionItem("PD") = "0" Then intReturnPageNum = 1

'Redirect
If blnCheckFirst Then	
	'Redirect to a page letting the user know their post is check first
	Response.Redirect("forum_posts.asp?TID=" & lngTopicID &	"&MF=Y&PID=" & lngMessageID & strQsSID3)
Else
	'Return	to the page showing the	posts
	Response.Redirect("forum_posts.asp?TID=" & lngTopicID &	"&PID=" & lngMessageID & strQsSID3)
End If
%>