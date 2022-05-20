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


'Dimension variables
Dim lngNumOfPosts		'Holds the number of posts a user has made
Dim lngNumOfPoints		'Holds the number of points the user has
Dim blnEmailNotify		'Set to	true if	the users want to be notified by e-mail	of a post
Dim blnEmailSent		'Set to	true if	the e-mail is sent
Dim strEmailSubject		'Holds the subject of the e-mail
Dim strMessage			'Holds the Users Message
Dim lngMessageID		'Holds the message ID number
Dim strMode			'Holds the mode	of the page so we know whether we are editing, updating, or new	topic
Dim lngTopicID			'Holds the topic ID number
Dim strSubject			'Holds the subject
Dim strPostDateTime		'Holds the current date	and time for the post
Dim strUserName			'Holds the username of the person we are going to email
Dim lngEmailUserID		'Holds the users ID of the person we are going to email
Dim strUserEmail		'Holds the users e-mail	address
Dim strEmailMessage		'Holds the body	of the e-mail
Dim blnSignature		'Holds wether a	signature is to	be shown or not
Dim intPriority			'Holds the priority of tipics
Dim strPostMessage		'Holds the post	to send	as mail	notify
Dim intReturnPageNum		'Holds the page	number to return to
Dim strForumName		'Holds the name	of the forum the message is being posted in
Dim intNumOfPostsInFiveMin	'Holds the number of posts the user has	made in	the last 5 minutes
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
Dim intNewGroupID		'Holds the new group ID for the poster
Dim strGuestName		'Holds the name of the guest if it is a guest posting
Dim intReplyCount		'Holds the number of threads in topic
Dim saryFileUploads		'Holds the names of the files uploaded
Dim objFSO			'Holds the file system object
Dim intLoop			'Loop counter
Dim strTopicIcon		'Holds the topic icon for the message
Dim dtmEventDate		'Holds the Calendar event date
Dim dtmEventDateEnd		'Holds the Calendar event date
Dim saryEmailNotify		'Holds the name of the person to email notify
Dim intEmailNotifyGroupID	'Email notify group ID
Dim intCurrentRecord		'Loop record count
Dim dtmAntiSpamTime		'The time and date for the antispam check
Dim objRegExp			'used for searches
Dim dtmLastForumPostDate	'Holds the date of the last forum post
Dim dtmLastTopicPostDate	'Holds the date of the last topic post


'Initalise variables
strPostDateTime	= internationalDateTime(Now())
dtmAntiSpamTime = Now()
intNumOfPostsInFiveMin = 0
lngPollID = 0
intReplyCount = 1
blnForumLocked = False
blnTopicLocked = False




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
'If the user is in a guest then get the Guest name they have entered
If lngLoggedInUserID = 2 Then strGuestName = Trim(Mid(Request.Form("Gname"), 1, 20))
	
'Read in event start date
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
'***	  Check CAPTCHA for Guest	***
'******************************************

'If this is a Guest poster and CAPTCHA is enabled check CAPTCHA is enetered correctly
If blnCAPTCHAsecurityImages AND lngLoggedInUserID = 2 Then
	
	'If CAPTCHA security image not enetered corectly kick the user
	If LCase(getSessionItem("SCS")) <> LCase(Trim(Request.Form("securityCode"))) OR getSessionItem("SCS") = "" Then 
		
		'Distroy session variable
		Call saveSessionItem("SCS", "")
		
		'Clean up
		Call closeDatabase()
	
		'Redirect
		If strMode = "new" Then
			Response.Redirect("not_posted.asp?mode=CAPTCHA" & strQsSID3)
		Else
			Response.Redirect("not_posted.asp?mode=CAPTCHA&TID=" & lngTopicID & strQsSID3)
		End If
		
	End If
	
	'Distroy session variable
	Call saveSessionItem("SCS", "")
End If





'******************************************
'***	     Get permissions	      *****
'******************************************


'If this is a new topic then only check the forum permissions
If strMode = "new" OR strMode = "poll" Then
	
	'As this is a new topic get the forum permissions
	strSQL = " " & _
	"SELECT" & strDBTop1 & " " & strDbTable & "Forum.Forum_name, " & strDbTable & "Forum.Password, " & strDbTable & "Forum.Forum_code, " & strDbTable & "Forum.Locked AS ForumLocked, " & strDbTable & "Forum.Password, " & strDbTable & "Permissions.* " & _
	"FROM " & strDbTable & "Forum" & strDBNoLock & ", " & strDbTable & "Permissions" & strDBNoLock & " " & _
	"WHERE " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Permissions.Forum_ID " & _
		"AND " & strDbTable & "Forum.Forum_ID = " & intForumID & " " & _
		"AND (" & strDbTable & "Permissions.Author_ID = " & lngLoggedInUserID & " OR " & strDbTable & "Permissions.Group_ID = " & intGroupID & ") " & _
	"ORDER BY " & strDbTable & "Permissions.Author_ID DESC" & strDBLimit1 & ";"
	
	'Query the database
	rsCommon.Open strSQL, adoCon

'Else if this is a reply in a topic check if the topic is locked also and get the forum ID from the topic that is being posted in
Else


	'As this is a reply in a topic get the forum permissions from the topic being posted in and also check if the topic is locked
	strSQL = " " & _
	"SELECT" & strDBTop1 & " " & strDbTable & "Forum.Forum_name, " & strDbTable & "Forum.Password, " & strDbTable & "Forum.Forum_code, " & strDbTable & "Forum.Locked AS ForumLocked, " & strDbTable & "Forum.Password, " & strDbTable & "Topic.Subject, " & strDbTable & "Topic.Locked AS TopicLocked, " & strDbTable & "Permissions.* " & _
	"FROM " & strDbTable & "Forum" & strDBNoLock & ", " & strDbTable & "Topic" & strDBNoLock & ", " & strDbTable & "Permissions" & strDBNoLock & " " & _
	"WHERE  " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Topic.Forum_ID " & _
		"AND " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Permissions.Forum_ID " & _
		"AND " & strDbTable & "Topic.Topic_ID = " & lngTopicID & " " & _
		"AND (" & strDbTable & "Permissions.Author_ID = " & lngLoggedInUserID & " OR " & strDbTable & "Permissions.Group_ID = " & intGroupID & ") " & _
	"ORDER BY " & strDbTable & "Permissions.Author_ID DESC" & strDBLimit1 & ";"
	
	'Query the database
	rsCommon.Open strSQL, adoCon
	
	'Read in some extra details as this is a reply
	If NOT rsCommon.EOF Then
		intForumID = CInt(rsCommon("Forum_ID"))
		strSubject = rsCommon("Subject")
		If blnAdmin = False Then blnTopicLocked = CBool(rsCommon("TopicLocked"))
	End If
End If


'Check the forum permissions
If NOT rsCommon.EOF Then

	strForumName = rsCommon("Forum_name")
	
	'See if the forum is locked if this is not the admin
	If blnAdmin = False Then blnForumLocked = CBool(rsCommon("ForumLocked"))

	'Read in the forum permissions
	blnRead = CBool(rsCommon("View_Forum"))
	blnPost = CBool(rsCommon("Post"))
	blnReply = CBool(rsCommon("Reply_posts"))
	blnEdit = CBool(rsCommon("Edit_posts"))
	blnPriority = CBool(rsCommon("Priority_posts"))
	blnPollCreate = CBool(rsCommon("Poll_create"))
	blnModerator = CBool(rsCommon("Moderate"))
	blnCheckFirst = CBool(rsCommon("Display_post"))
	blnEvents = CBool(rsCommon("Calendar_event"))
	
	'If the user has no read rights then kick them
	If blnRead = False Then
		
		'Reset Server Objects
		rsCommon.Close
		Call closeDatabase()

		'Kick the user
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
		
	'If this in not an admin or moderator set the priority to 0
	If (blnAdmin = false OR blnModerator = false) AND blnPriority = false Then intPriority = 0




'Else nothing returned from db kicjk user
Else
	'Reset Server Objects
	rsCommon.Close
	Call closeDatabase()

	'Kick the user
	Response.Redirect("insufficient_permission.asp" & strQsSID1)

End If

'Clean up
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

'Remove start of character put in for firefox bug
strMessage = Replace(strMessage, chr(2), "")
strMessage = Replace(strMessage, chr(127), "")

'Remove any tabs from subject
strSubject = Replace(strSubject, vbTab, " ")

'Fix for UK Pound Sterling displaying as Â£
strSubject = Replace(strSubject, Chr(194) & Chr(163), "£", 1, -1, 1)
 

'If there is no	subject	or message then	don't post the message as won't	be able	to link	to it
If strSubject =	"" AND (strMode	= "new" OR strMode = "poll") Then strReturnCode = "noSubject"
If Trim(strMessage) = "" OR Trim(strMessage) = "<P>&nbsp;</P>" OR Trim(strMessage) = "<br>" OR Trim(strMessage) = "<br>" & vbCrLf OR Trim(strMessage) = vbCrLf Then strReturnCode = "noSubject"


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
If strMode = "poll" AND	blnPollCreate Then

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

		'If there is nothing in	the poll selection then	jump out the loop
		If saryPollChoice(intPollChoice) = "" Then

			'ReDimension the array for the correct number of choices
			ReDim Preserve saryPollChoice(intPollChoice - 1)

			'Exit loop
			Exit For
		End If

		'Clean up input
		saryPollChoice(intPollChoice) =	removeAllTags(saryPollChoice(intPollChoice))
	Next
End If





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




'******************************************
'***	      Anti-spam	Check		***
'******************************************

'Initalise the SQL string with a query to read in the last post	from the database
strSQL = "SELECT "
If strDatabaseType = "SQLServer" OR strDatabaseType = "Access" Then
	strSQL = strSQL & "TOP 15"
End If
strSQL = strSQL & " " & strDbTable & "Thread.Message, " & strDbTable & "Thread.Author_ID, " & strDbTable & "Thread.Message_date " & _
"FROM " & strDbTable & "Thread " & strDBNoLock & " " & _
"ORDER BY " & strDbTable & "Thread.Message_date DESC"
If strDatabaseType = "mySQL" Then
	strSQL = strSQL & " LIMIT 15"
End If
strSQL = strSQL & ";"

'Open the recordset
rsCommon.Open strSQL, adoCon

'If there is a post returned by	the recorset then check	it's not already posted	and for	spammers
If NOT rsCommon.EOF Then
	'Check the last	message	posted is not the same as the new one
	If (rsCommon("Message")	= strMessage) Then

		'Set the return	code
		strReturnCode =	"posted"
	End If

	'Check the user	hasn't posted in the last limit	set for	secounds and not more than 5 times in the last spam time limit set for minutes
	Do While NOT rsCommon.EOF AND blnAdmin = False AND lngLoggedInUserID <> 2
		
		'Check that the spam time to test is not negative (this is for daylight saving and if the forum is moved to a new server with negartive time)
		If DateDiff("h", rsCommon("Message_date"), dtmAntiSpamTime) < 0 Then 
			'If the antispam time is negative hours then add the negative hours to get postive hoyrs
			dtmAntiSpamTime = DateAdd("h", CInt(Replace(CStr(DateDiff("h", rsCommon("Message_date"), dtmAntiSpamTime)), "-", "")), dtmAntiSpamTime)
		End If

		'Check the user	hasn't posted in the last spam time limit set for seconds
		If rsCommon("Author_ID") = lngLoggedInUserID AND DateDiff("s", rsCommon("Message_date"), dtmAntiSpamTime)	< intSpamTimeLimitSeconds AND intSpamTimeLimitSeconds <> 0 Then

			'Set the return	code
			strReturnCode =	"maxS"
		End If

		'Check that the	user hasn't posted 5 posts in the spam time limit set for minutes
		If rsCommon("Author_ID") = lngLoggedInUserID AND DateDiff("n", rsCommon("Message_date"), dtmAntiSpamTime)	< intSpamTimeLimitMinutes AND intSpamTimeLimitMinutes <> 0 Then

			'Add 1 to the number of	posts in the last 5 minutes
			intNumOfPostsInFiveMin = intNumOfPostsInFiveMin	+ 1

			'If the	number of posts	is more	than 3 then set	the return code
			If intNumOfPostsInFiveMin = 5 Then

				'Set the return	code
				strReturnCode =	"maxM"
			End If
		End If

		'Move to the next post
		rsCommon.MoveNext
	Loop
End If

'Clean up
rsCommon.Close













'**********************************************
'***  If input problems	send to	error page  ***
'**********************************************

'If there is a return code then	this post is not valid so redirect to error page
If strReturnCode <> "" Then

	'Clean up
	Call closeDatabase()

	'Redirect to error page
	Response.Redirect("not_posted.asp?mode=" & strReturnCode & strQsSID3)
End If




'********************************************
'***		  Save new Poll		  ***
'********************************************

'If this is a poll then save the poll to the database
If strMode = "poll" AND	blnPollCreate Then

	'********************************************
	'***	     Save poll question		  ***
	'********************************************

	'Initalise the SQL string with a query to get the poll last poll details to get the poll ID number in next (use nolock as this is a new insert so a dirty read is OK)
	strSQL = "SELECT" & strDBTop1 & " " & strDbTable & "Poll.* " & _
	"FROM " & strDbTable & "Poll" & strRowLock & " " & _
	"ORDER BY " & strDbTable & "Poll.Poll_ID DESC" & strDBLimit1 & ";"

	With rsCommon
		'Set the cursor	type property of the record set	to Forward only
		.CursorType = 0

		'Set the Lock Type for the records so that the record set is only locked when it is updated
		.LockType = 3

		'Open the author table
		.Open strSQL, adoCon

		'Insert	the new	poll question in the recordset
		.AddNew

		'Update	recordset
		.Fields("Poll_question") = strPollQuestion
		.Fields("Multiple_votes") = blnMultipleVotes
		.Fields("Reply") = blnPollReply

		'Update	the database with the new poll question
		.Update

		'Re-run	the Query once the database has	been updated to get the poll's ID number
		.Requery

		'Read in the new poll's ID number
		lngPollID = CLng(rsCommon("Poll_ID"))

		'Clean up
		.Close
	End With


	'********************************************
	'***	      Save poll	choices		  ***
	'********************************************

	'Initalise the SQL string with a query to get the choice (use nolock as this is a new insert so a dirty read is OK)
	strSQL = "SELECT " & strDbTable & "PollChoice.* " & _
	"FROM " & strDbTable & "PollChoice" & strRowLock & " " & _
	"WHERE " & strDbTable & "PollChoice.Poll_ID = 0;"

	With rsCommon
		'Set the cursor	type property of the record set	to Forward only
		.CursorType = 0

		'Set the Lock Type for the records so that the record set is only locked when it is updated
		.LockType = 3

		'Open the author table
		.Open strSQL, adoCon

		'Add the new poll choices to recordset
		For intPollChoice = 1 To UBound(saryPollChoice)
		
			'Insert	the new	poll choices in	the recordset
			.AddNew

			'Update	recordset
			.Fields("Poll_ID") = lngPollID
			.Fields("Choice") = saryPollChoice(intPollChoice)
		Next

		'Update	the database with the new poll choices
		.Update

		'Clean up
		.Close
	End With

	'Change	the mode to new	to save	the new	polls post message
	strMode = "new"
End If





'******************************************
'***	 Save new topic	subject		***
'******************************************

'If this is a new topic	then save the new subject heading and read back	the new	topic ID number
If strMode = "new" AND (blnPost OR blnPollCreate OR (blnAdmin OR blnModerator)) Then

	'Initalise the SQL string with a query to get the Topic	details
	strSQL = "SELECT" & strDBTop1 & " " & strDbTable & "Topic.Topic_ID, " & strDbTable & "Topic.Forum_ID, " & strDbTable & "Topic.Poll_ID, " & strDbTable & "Topic.Icon, " & strDbTable & "Topic.Subject, " & strDbTable & "Topic.Priority, " & strDbTable & "Topic.Hide, " & strDbTable & "Topic.Event_date, " & strDbTable & "Topic.Event_date_end, " & strDbTable & "Topic.Start_Thread_ID, " & strDbTable & "Topic.Last_Thread_ID " & _
	"FROM " & strDbTable & "Topic" & strRowLock & " " & _
	"WHERE " & strDbTable & "Topic.Forum_ID = " & intForumID & " "  & _
	"ORDER BY " & strDbTable & "Topic.Topic_ID DESC" & strDBLimit1 & ";"

	With rsCommon
		'Set the cursor	type property of the record set	to Forward only
		.CursorType = 0

		'Set the Lock Type for the records so that the record set is only locked when it is updated
		.LockType = 3

		'Open the author table
		.Open strSQL, adoCon
		
		'Set error trapping
		On Error Resume Next

		'Insert	the new	topic details in the recordset
		.AddNew

		'Update	recordset
		.Fields("Forum_ID") = intForumID
		.Fields("Poll_ID") = lngPollID
		If blnTopicIcon Then .Fields("Icon") = strTopicIcon
		.Fields("Subject") = strSubject
		.Fields("Priority") = intPriority
		.Fields("Hide") = blnCheckFirst
		'If Calendar events allowed save 'em
		If blnCalendar AND blnEvents Then .Fields("Event_date") = dtmEventDate
		If blnCalendar AND blnEvents Then .Fields("Event_date_end") = dtmEventDateEnd
		
		'Create a dummy random negative number for start and end Thread ID, this is set correctly later after the post is saved and the real Thread ID is available	
		Randomize
		.Fields("Start_Thread_ID") = - CInt(Rnd * 1000)
		.Fields("Last_Thread_ID") = - CInt(Rnd * 1000)
		


		'Update	the database with the new topic	details
		.Update
		
		'If an error has occurred write an error to the page
		If Err.Number <> 0 Then	Call errorMsg("An error has occurred while writing to the database.", "save_new_topic_data", "new_post.asp")
				
		'Disable error trapping
		On Error goto 0

		'Re-run	the Query once the database has	been updated
		.Requery

		'Read in the new topic's ID number
		lngTopicID = CLng(rsCommon("Topic_ID"))

		'Set the rerun page properties
		intReturnPageNum = 1

		'Clean up
		.Close
	End With
	
	'If logging enabled log the new topic has been created
	If blnLoggingEnabled AND blnCreatePostLogging Then Call logAction(strLoggedInUsername, "Created New Topic '" & decodeString(strSubject) & "' - TopicID " & lngTopicID)
End If




'******************************************
'***	  Process New Post		***
'******************************************

'This is a new post so save the new post to the database
If ((strMode = "new" AND (blnPost OR blnPollCreate)) OR (blnReply)) OR (blnAdmin OR blnModerator) Then


	'******************************************
	'***	       Save New	Post		***
	'******************************************

	'Initalise the strSQL variable with an SQL statement to	query the database get the message details
	'Don't use no lock as we need a clean read when getting the thread ID
	strSQL = "SELECT" & strDBTop1 & " " & strDbTable & "Thread.Thread_ID, " & strDbTable & "Thread.Topic_ID,  " & strDbTable & "Thread.Author_ID, " & strDbTable & "Thread.Message, " & strDbTable & "Thread.Message_date, " & strDbTable & "Thread.Show_signature, " & strDbTable & "Thread.IP_addr, " & strDbTable & "Thread.Hide " & _
	"FROM	" & strDbTable & "Thread" & strRowLock & " " & _
	"ORDER BY " & strDbTable & "Thread.Thread_ID DESC" & strDBLimit1 & ";"

	With rsCommon
		'Set the cursor	type property of the record set	to Forward only
		.CursorType = 0

		'Set the Lock Type for the records so that the record set is only locked when it is updated
		.LockType = 3

		'Open the threads table
		.Open strSQL, adoCon
		
		'Set error trapping
		On Error Resume Next

		'Insert	the new	Thread details in the recordset
		.AddNew

		.Fields("Topic_ID") = lngTopicID
		.Fields("Author_ID") = lngLoggedInUserID
		.Fields("Message") = strMessage
		.Fields("Message_date")	= strPostDateTime
		.Fields("Show_signature") = blnSignature
		.Fields("IP_addr") = getIP()
		.Fields("Hide") = blnCheckFirst
		
		
		'Update	the database with the new Thread
		.Update
		
		'If an error has occurred write an error to the page
		If Err.Number <> 0 Then	Call errorMsg("An error has occurred while writing to the database.", "save_new_post_data", "new_post.asp")
				
		'Disable error trapping
		On Error goto 0

		'Requery cuase Access is so slow (needed to get	accurate post count)
		.Requery

		'Read in the thread ID for the guest posting
		lngMessageID = CLng(rsCommon("Thread_ID"))

		'Clean up
		.Close
	End With

	'If logging enabled log the new topic has been created
	If blnLoggingEnabled AND blnCreatePostLogging AND strMode = "reply"  Then Call logAction(strLoggedInUsername, "Posted Reply in Topic '" & decodeString(strSubject) & "' - PostID " & lngMessageID)


	'******************************************
	'***	 Update	Topic Last Post	Datails	***
	'******************************************
	
	'This is a new topic so place the start and last post author ID in the topic table (don't update the no. of replies)
	If strMode = "new" Then
		
		'Initalise the SQL string with an SQL update command to	update the last author
		strSQL = "UPDATE " & strDbTable & "Topic " & strRowLock & " " & _
		"SET " & strDbTable & "Topic.Start_Thread_ID = " & lngMessageID & ", " & _
			strDbTable & "Topic.Last_Thread_ID = " & lngMessageID & " " & _
		"WHERE " & strDbTable & "Topic.Topic_ID = " & lngTopicID & ";"
	
		'Write the updated date	of last	post to	the database
		adoCon.Execute(strSQL)
	
	
	'If the post is displayed update the no. of replies and last author date
	ElseIf blnCheckFirst = false Then
		
		'Update the stats fro this topic
		Call updateTopicStats(lngTopicID)
	End If
	
	
	
	'******************************************
	'***  Update The Forum stats		***
	'******************************************

	'Update	the forum stats for main page (no. of posts+topics, last post author, and last post date)
	If blnCheckFirst = false Then Call updateForumStats(intForumID)




	'******************************************
	'***	 Save the guest username	***
	'******************************************

	'If this is a guest that is posting then save there name to the db
	If lngLoggedInUserID = 2 AND strGuestName <> "" Then
		'Initalise the SQL string with an SQL update command to	update the date	of the last post in the	Topic table
		strSQL = "INSERT INTO " & strDbTable & "GuestName (" & _
		"Name, " & _
		"Thread_ID " & _
		") " & _
		"VALUES " & _
		"('" & strGuestName & "', " & _
		"'" & lngMessageID & "' " & _
		")"

		'Write the updated date	of last	post to	the database
		adoCon.Execute(strSQL)
	End If




	'*****************************************************
	'*** Update Author No. of Posts/Points/Active Time ***
	'*****************************************************

	'Initalise the strSQL variable with an SQL statement to	query the database to get the number of	posts the user has made
	strSQL = "SELECT " & strDbTable & "Author.No_of_posts, " & strDbTable & "Author.Points, " & strDbTable & "Group.Special_rank, " & strDbTable & "Group.Ladder_ID " & _
	"FROM	" & strDbTable & "Author " & strDBNoLock & ", " & strDbTable & "Group " & strDBNoLock & " " & _
	"WHERE " & strDbTable & "Author.Group_ID = " & strDbTable & "Group.Group_ID " & _
		"AND " & strDbTable & "Author.Author_ID = " & lngLoggedInUserID & ";"
		
	'Query the database
	rsCommon.Open strSQL, adoCon	
		
	'If there is a record returned by the database then read in the	no of posts and	increment it by	1
	If NOT rsCommon.EOF Then
		
		'Read in the no	of posts the user has made and username
		If isNull(rsCommon("No_of_posts")) Then lngNumOfPosts = 0 Else lngNumOfPosts = CLng(rsCommon("No_of_posts"))
		If isNull(rsCommon("Points")) Then lngNumOfPoints = 0 Else lngNumOfPoints = CLng(rsCommon("Points"))
		
		'Inrement the number of	posts by 1
		lngNumOfPosts =	lngNumOfPosts +	1
		
		'Inrement the number of points by the correct amount
		If strMode = "new" OR strMode = "poll" Then
			lngNumOfPoints = lngNumOfPoints + intPointsTopic
		Else
			lngNumOfPoints = lngNumOfPoints + intPointsReply
		End If
			
		'Initalise the SQL string with an SQL update command to	update the number of posts the user has	made
		strSQL = "UPDATE " & strDbTable & "Author " & strRowLock & " " & _
		"SET " & strDbTable & "Author.No_of_posts = " & lngNumOfPosts & ", " & _
			strDbTable & "Author.Points = " & lngNumOfPoints & ", " & _
			strDbTable & "Author.Last_visit = " & formatDbDate(Now()) & _
		"WHERE " & strDbTable & "Author.Author_ID = " & lngLoggedInUserID & ";"
				
		'Write the updated number of posts to the database
		adoCon.Execute(strSQL)
	End If
	
	
	
	
	'******************************************
	'***	    Update Ladder Group	        ***
	'******************************************
		
	'See if the user is a member of a ladder group and if so update their group if they have enough posts
	
	'If there is a record returned by the database then see if it is a group that needs updating
	If NOT rsCommon.EOF Then
	
		'If a ladder group then see if the group needs updating
		If CBool(rsCommon("Special_rank")) = False Then
			
			Dim intLadderGroup
			
			'Read in the ladder group
			intLadderGroup = CInt(rsCommon("Ladder_ID"))
	
			'Clean up
			rsCommon.Close
	
			'Initlise variables
			intNewGroupID = intGroupID
	
			'Get the rank group the member should be part of
			'Initalise the strSQL variable with an SQL statement to	query the database to get the number of	posts the user has made
			strSQL = "SELECT" & strDBTop1 & " " & strDbTable & "Group.Group_ID "  & _
			"FROM " & strDbTable & "Group " & strDBNoLock & " "  & _
			"WHERE (" & strDbTable & "Group.Minimum_posts <= " & lngNumOfPoints & ") " & _
				"AND (" & strDbTable & "Group.Minimum_posts >= 0) "  & _
				"AND (" & strDbTable & "Group.Ladder_ID = " & intLadderGroup & ") " & _
			"ORDER BY " & strDbTable & "Group.Minimum_posts DESC" & strDBLimit1 & ";"
	
			'Query the database
			rsCommon.Open strSQL, adoCon
	
	
			'Get the new Group ID
			If NOT rsCommon.EOF Then intNewGroupID = CInt(rsCommon("Group_ID"))
	
	
			'If the group ID is different to the present group one then update it
			If intGroupID <> intNewGroupID Then
	
				'Initalise the SQL string with an SQL update command to	update group ID of the author
				strSQL = "UPDATE " & strDbTable & "Author " & strRowLock & " " & _
				"SET " & strDbTable & "Author.Group_ID = " & intNewGroupID & " " & _
				"WHERE " & strDbTable & "Author.Author_ID = " & lngLoggedInUserID & ";"
	
				'Write the updated number of posts to the database
				adoCon.Execute(strSQL)
			End If
		End If
	End If
	
	'Close the recordset
	rsCommon.Close




	'******************************************
	'***	   Send	Email Notification	 **
	'******************************************

	If blnEmail Then

		'If post approval is enabled only send emails to the moderators and not anyone else
		If blnCheckFirst Then

			'******************************************
			'*** Admin/Moderator Email Notification ***
			'******************************************
			
			
			'Set the e-mail	subject
			strEmailSubject	= strMainForumName & " " & strTxtPendingApproval & " : " & decodeString(strSubject)
	
			'If we are to send an e-mail notification and send the post with the e-mail then format	the post for the e-mail
			If blnSendPost Then
				
				'Format	the post to be sent with the e-mail
				strPostMessage = "<br /><strong>" & strTxtForum & ":</strong> " & strForumName & _
				"<br /><strong>" & strTxtTopic & ":</strong> " & strSubject & _
				"<br /><strong>" & strTxtPostedBy & ":</strong> " & strLoggedInUsername & "<br /><br />" & strMessage
	
				'Change	the path to the	emotion	symbols	to include the path to the images
				strPostMessage = Replace(strPostMessage, "src=""smileys/", "src=""" & strForumPath & "smileys/", 1, -1, 1)
			End If
			
			'Initalise the strSQL variable with an SQL statement to	query the database get the details for the email
			strSQL = "SELECT DISTINCT " & strDbTable & "EmailNotify.Author_ID, " & strDbTable & "Author.Username,	" & strDbTable & "Author.Author_email, " & strDbTable & "Author.Group_ID "  & _
			"FROM	" & strDbTable & "Author" & strDBNoLock & ", " & strDbTable & "EmailNotify" & strDBNoLock & ", " & strDbTable & "Topic" & strDBNoLock & " "  & _
			"WHERE " & strDbTable & "Author.Author_ID = " & strDbTable & "EmailNotify.Author_ID "  & _
				"AND (" & strDbTable & "EmailNotify.Forum_ID = " & intForumID & " OR " & strDbTable & "EmailNotify.Topic_ID = " & lngTopicID & ") "  & _
				"AND (" & strDbTable & "Author.Group_ID " & _
					"IN (" & _
						"SELECT " & strDbTable & "Permissions.Group_ID " & _
						"FROM " & strDbTable & "Permissions " & strDBNoLock & " " & _
						"WHERE (" & strDbTable & "Permissions.Moderate = " & strDBTrue & ") " & _
						") " & _
					"OR " & strDbTable & "Author.Author_ID " & _
						"IN (" & _
							"SELECT " & strDbTable & "Permissions.Author_ID " & _
							"FROM " & strDbTable & "Permissions " & strDBNoLock & " " & _
							"WHERE (" & strDbTable & "Permissions.Moderate = " & strDBTrue & ") " & _
							") " & _
					"OR " & _
						strDbTable & "Author.Group_ID = 1 " & _
				")" & _
				"AND " & strDbTable & "Author.Author_email Is Not Null "  & _
				"AND " & strDbTable & "Author.Banned = " & strDBFalse & " " & _
				"AND " & strDbTable & "Author.Active = " & strDBTrue & " " & _
				"AND " & strDbTable & "Author.Last_visit > " & formatDbDate(DateAdd("yyyy", -1, now())) & ";"

			
			'Query the database
			rsCommon.Open strSQL, adoCon
			
		
			
		
		
		'Else send out email notifications to all subscribers
		Else
			
			'******************************************
			'*** All Subscribers Email Notification ***
			'******************************************
			
			'Set the e-mail	subject
			strEmailSubject	= decodeString(strSubject)
	
			'If we are to send an e-mail notification and send the post with the e-mail then format	the post for the e-mail
			If blnSendPost Then
	
				'Format	the post to be sent with the e-mail
				strPostMessage = "<br /><strong>" & strTxtForum & ":</strong> " & strForumName & _
				"<br /><strong>" & strTxtTopic & ":</strong> " & strSubject & _
				"<br /><strong>" & strTxtPostedBy & ":</strong> " & strLoggedInUsername & "<br /><br />" & strMessage
	
				'Change	the path to the	emotion	symbols	to include the path to the images
				strPostMessage = Replace(strPostMessage, "src=""smileys/", "src=""" & strForumPath & "smileys/", 1, -1, 1)
			End If
			
			
			
			'Updated to only send one notification per forum or topic since members last visit
			If blnEmailNotificationSendAll = False Then
			
				'Set the last post date/time the same as message date/time incase it is a new forum or topic
				dtmLastForumPostDate = strPostDateTime
				dtmLastTopicPostDate = strPostDateTime
				
				
				'Get the last forum post date 
				strSQL = "SELECT" & strDBTop1 & " " & strDbTable & "Thread.Message_date " & _
				"FROM " & strDbTable & "Topic" & strDBNoLock & ", " & strDbTable & "Thread" & strDBNoLock & " " & _
				"WHERE " & strDbTable & "Topic.Topic_ID = " & strDbTable & "Thread.Topic_ID " & _
					"AND " & strDbTable & "Topic.Forum_ID = " & intForumID & " " & _
					"AND " & strDbTable & "Thread.Message_date < " & formatDbDate(strPostDateTime) & " " & _
				"ORDER BY " & strDbTable & "Thread.Message_date DESC" & strDBLimit1 & ";"
				
				'Query the database
				rsCommon.Open strSQL, adoCon
				
				'Read in the last forum post date
				If NOT rsCommon.EOF Then dtmLastForumPostDate = rsCommon("Message_date")
				
				'Close rs
				rsCommon.Close
				
				
				'Get the last topic post date
				strSQL = "SELECT" & strDBTop1 & " " & strDbTable & "Thread.Message_date " & _
				"FROM " & strDbTable & "Thread" & strDBNoLock & " " & _
				"WHERE " & strDbTable & "Thread.Topic_ID = " & lngTopicID & " " & _
					"AND " & strDbTable & "Thread.Message_date < " & formatDbDate(strPostDateTime) & " " & _
				"ORDER BY " & strDbTable & "Thread.Message_date DESC" & strDBLimit1 & ";"
				
				'Query the database
				rsCommon.Open strSQL, adoCon
				
				'Read in the last forum post date
				If NOT rsCommon.EOF Then dtmLastTopicPostDate = rsCommon("Message_date")
				
				'Close rs
				rsCommon.Close
			
			'Else set date of the last post to the year that Web Wiz was launched
			Else
				dtmLastForumPostDate = "2001-01-01"
				dtmLastTopicPostDate = "2001-01-01"
			
			End If
			
			
			'Initalise the strSQL variable with an SQL statement to	query the database get the details for the email
			strSQL = "SELECT DISTINCT " & strDbTable & "EmailNotify.Author_ID, " & strDbTable & "Author.Username,	" & strDbTable & "Author.Author_email, " & strDbTable & "Author.Group_ID "  & _
			"FROM	" & strDbTable & "Author" & strDBNoLock & ", " & strDbTable & "EmailNotify" & strDBNoLock & " "  & _
			"WHERE " & strDbTable & "Author.Author_ID = " & strDbTable & "EmailNotify.Author_ID "  & _
				"AND ((" & strDbTable & "EmailNotify.Forum_ID = " & intForumID & " AND " & strDbTable & "Author.Last_visit > " & formatDbDate(dtmLastForumPostDate) & ") "  & _
					"OR  (" & strDbTable & "EmailNotify.Topic_ID = " & lngTopicID & " AND " & strDbTable & "Author.Last_visit > " & formatDbDate(dtmLastTopicPostDate) & ")) "  & _
				"AND " & strDbTable & "Author.Author_email Is Not Null "  & _
				"AND " & strDbTable & "Author.Author_email <> '' "  & _
				"AND " & strDbTable & "Author.Banned = " & strDBFalse & " " & _
				"AND " & strDbTable & "Author.Active = " & strDBTrue & " " & _
				"AND " & strDbTable & "Author.Last_visit > " & formatDbDate(DateAdd("yyyy", -1, now())) & ";"
				
			'Query the database
			rsCommon.Open strSQL, adoCon
			
			
		
			
		End If
		
		'Place RS into array
		If NOT rsCommon.EOF Then saryEmailNotify = rsCommon.GetRows()
		
		'Close RS
		rsCommon.Close
		
		
		'*******************************************
		'***	   Send	Email Notification	 ***
		'*******************************************

		'If a record is	returned by the	recordset then read in the details and send the	email
		If isArray(saryEmailNotify) Then
			
			'Loop through sending email notifications
			Do While intCurrentRecord <= Ubound(saryEmailNotify,2)

				'Read in the details from the recordset	for the	e-mail
				lngEmailUserID = CLng(saryEmailNotify(0,intCurrentRecord))
				strUserName = saryEmailNotify(1,intCurrentRecord)
				strUserEmail = saryEmailNotify(2,intCurrentRecord)
				intEmailNotifyGroupID = CInt(saryEmailNotify(3,intCurrentRecord))
				
				
				'Check the email recepient has permission in this forum
				strSQL = "SELECT " & strDbTable & "Permissions.View_Forum " & _
				"FROM " & strDbTable & "Permissions" & strDBNoLock & " " & _
				"WHERE (" & strDbTable & "Permissions.Author_ID = " & lngEmailUserID & " OR " & strDbTable & "Permissions.Group_ID = " & intEmailNotifyGroupID & ") " & _
				"AND " & strDbTable & "Permissions.View_Forum = " & strDBTrue & ";"
				
				'Query the database
				rsCommon.Open strSQL, adoCon
				
				
				'If no record returned then user is not allowed in this forum so don't send email notifications
				If rsCommon.EOF Then
					
					'User doesn't have email notifications so delete any they have in this forum or topic
					strSQL = "DELETE FROM " & strDbTable & "EmailNotify " & _
					"WHERE " & strDbTable & "EmailNotify.Author_ID = " & lngEmailUserID & " " & _
						"AND (" & strDbTable & "EmailNotify.Forum_ID = " & intForumID & " OR " & strDbTable & "EmailNotify.Topic_ID = " & lngTopicID & ");"
					
					'Execute query
					adoCon.Execute(strSQL)
				
	
				'If the	user wants to be e-mailed and the user has enetered their e-mail and they are not the original topic writter then send an e-mail
				ElseIf lngEmailUserID <> lngLoggedInUserID AND Trim(strUserEmail) <> "" Then
					
					'Initailise the	e-mail body variable with the body of the e-mail
					strEmailMessage	= strTxtHi & " " & decodeString(strUserName) & ","
					If blnCheckFirst Then
						strEmailMessage	= strEmailMessage & "<br /><br />" & strTxtEmailAMeesageHasBeenPosted & " " & strMainForumName & " " & strTxtThatRequiresApproval
					Else
						strEmailMessage	= strEmailMessage & "<br /><br />" & strTxtEmailAMeesageHasBeenPosted & " " & strMainForumName & " " & strTxtThatYouAskedKeepAnEyeOn
					End If
					strEmailMessage	= strEmailMessage & _
					"<br /><br />" & strTxtEmailClickOnLinkBelowToView & " : -" & _
					"<br /><a href=""" & strForumPath & "forum_posts.asp?TID=" & lngTopicID & "&PID=" & lngMessageID & "&Type=EmailNotification#" & lngMessageID & """>" & strForumPath & "forum_posts.asp?TID=" & lngTopicID & "&PID=" & lngMessageID & "&Type=EmailNotification#" & lngMessageID & "</a>" & _
					"<br /><br />" & strTxtClickTheLinkBelowToUnsubscribe & " :	-" & _
					"<br /><a href=""" & strForumPath & "email_notify.asp?TID=" & lngTopicID & "&FID=" & intForumID & "&M=Unsubscribe"">" & strForumPath & "email_notify.asp?TID=" & lngTopicID & "&FID=" & intForumID & "&M=Unsubscribe</a>"
					If blnCheckFirst = False Then strEmailMessage = strEmailMessage & "<br /><br />" & strTxtThereMayAlsoBeOtherMessagesPostedOn & "."
					
					'If we are to send the post then attach	it as well
					If blnSendPost = True Then strEmailMessage = strEmailMessage & "<br /><br /><hr />" & strPostMessage & "<br />"
			
	
					'Call the function to send the e-mail
					blnEmailSent = SendMail(strEmailMessage, decodeString(strUserName), decodeString(strUserEmail),	strWebsiteName, decodeString(strForumEmailAddress), decodeString(strEmailSubject), strMailComponent, true)
				End If
	
				'Move to the next record in the	recordset
				intCurrentRecord = intCurrentRecord + 1
				
				'Close the recordset
				rsCommon.Close
			Loop
		End If
		
	End If
End If





'**********************************************************
'***	     Update Email Notify if this is a reply	***
'**********************************************************

'Delete	or Save	email notification for the user, if email notify is enabled

If blnEmail = True Then

	'Initalise the SQL string with a query to get the email	notify details
	strSQL = "SELECT " & strDbTable & "EmailNotify.*  " & _
	"FROM	" & strDbTable & "EmailNotify" & strRowLock & " " & _
	"WHERE " & strDbTable & "EmailNotify.Author_ID = " & lngLoggedInUserID & " "  & _
		"AND " & strDbTable & "EmailNotify.Topic_ID = " & lngTopicID & ";"

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
	
	'If this is a Guest then send them to a page letting them know there post needs to be moderated before being displayed
	If intGroupID = 2 Then
		Response.Redirect("moderated_post.asp?FID=" & intForumID & "&TID=" & lngTopicID & "&PN=" & intReturnPageNum & strQsSID3 & "&M=" & strMode)
	
	'Redirect to a page letting the user know their post is check first
	Else
		Response.Redirect("forum_posts.asp?TID=" & lngTopicID &	"&MF=Y&PID=" & lngMessageID & strQsSID3 & "&#" & lngMessageID)
	End If
Else
	'Return	to the page showing the	posts 
	Response.Redirect("forum_posts.asp?TID=" & lngTopicID &	"&PID=" & lngMessageID & strQsSID3 & "&#" & lngMessageID)
End If
%>