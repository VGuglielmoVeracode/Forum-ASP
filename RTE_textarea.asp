<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="language_files/pm_language_file_inc.asp" -->
<!--#include file="functions/functions_date_time_format.asp" -->
<!--#include file="functions/functions_edit_post.asp" -->
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

'Dimension variables
Dim strMode		'Holds the mode of the page
Dim strMessage		'Holds the message to be edited
Dim lngPostID		'Holds the post ID number
Dim strQuoteUsername	'Holds the quoters username
Dim strQuoteMessage	'Holds the message to be quoted
Dim lngQuoteUserID	'Holds the quoters user ID
Dim dtmReplyPMDate	'Holds the PM date


'Read in the message ID number to edit
strMode = Request.QueryString("mode")
lngPostID = LngC(Request.QueryString("ID"))





'If the message is to be edited then read in the message from the database
If strMode = "edit" or strMode="editTopic" OR strMode = "editPoll" Then
	
	'Initalise the strSQL variable with an SQL statement to query the database get the message details
	strSQL = "SELECT " & strDbTable & "Thread.Message, " & strDbTable & "Forum.Forum_ID " & _
	"FROM " & strDbTable & "Forum" & strDBNoLock & ", " & strDbTable & "Topic" & strDBNoLock & ", " & strDbTable & "Thread" & strDBNoLock & " " & _
	"WHERE " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Topic.Forum_ID " & _
		"AND " & strDbTable & "Topic.Topic_ID = " & strDbTable & "Thread.Topic_ID " & _
		"AND " & strDbTable & "Thread.Thread_ID = " & lngPostID & ";"
	
	'Query the database
	rsCommon.Open strSQL, adoCon 
	
	'Read in the details from the recordset
	strMessage = rsCommon("Message")
	intForumID = CInt(rsCommon("Forum_ID"))
	
	'Clean up
	rsCommon.Close



'If the message is to have a quote from someone else then read in there message
ElseIf strMode = "quote" Then
	
	
	'Initalise the strSQL variable with an SQL statement to get the message to be quoted
	strSQL = "SELECT " & strDbTable & "Topic.Forum_ID, " & strDbTable & "Thread.Author_ID, " & strDbTable & "Thread.Message, " & strDbTable & "Author.Username, " & strDbTable & "GuestName.Name " & _
	"FROM (" & strDbTable & "Author" & strDBNoLock & " INNER JOIN (" & strDbTable & "Topic" & strDBNoLock & " INNER JOIN " & strDbTable & "Thread" & strDBNoLock & " ON " & strDbTable & "Topic.Topic_ID = " & strDbTable & "Thread.Topic_ID) ON " & strDbTable & "Author.Author_ID = " & strDbTable & "Thread.Author_ID) LEFT JOIN " & strDbTable & "GuestName" & strDBNoLock & " ON " & strDbTable & "Thread.Thread_ID = " & strDbTable & "GuestName.Thread_ID "  & _
	"WHERE " & strDbTable & "Thread.Thread_ID = " & lngPostID & ";"
		
	'Query the database
	rsCommon.Open strSQL, adoCon 
		
		
	'Read in the details from the recordset
	intForumID = CInt(rsCommon("Forum_ID"))
	lngQuoteUserID = CLng(rsCommon("Author_ID"))
	strQuoteUsername = rsCommon("Username")
	strQuoteMessage = rsCommon("Message")
	
	'If the post being quoted is written by a guest see if they have a name
	If lngQuoteUserID = 2 Then strQuoteUsername = rsCommon("Name")
	
	'Clean up
	rsCommon.Close
	
	
	'Build up the quoted thread post
	strMessage = "[QUOTE=" & strQuoteUsername & "]"
	
	'Read in the quoted thread from the recordset
	strMessage = strMessage & strQuoteMessage
	strMessage = strMessage & "[/QUOTE]"




'Else if this is a reply and we are going from a quick reply to full display so
ElseIf InStr(strMode, "QuickToFull") AND Session("Message") <> "" Then
	strMessage = Session("Message")
	Session("Message") = Null 


'If a private message read in the message again if the user has returned to ammend after getting username wrong	
ElseIf strMode = "PM" AND Session("PmMessage") <> "" Then
	strMessage = Session("PmMessage")
	Session("PmMessage") = Null 
End If




'If we are replying to a private message then format it
If strMode = "PM" AND NOT lngPostID = 0 Then
	
	'Initlise the sql statement
	strSQL = "SELECT " & strDbTable & "PMMessage.*, " & strDbTable & "Author.Username " & _
	"FROM " & strDbTable & "PMMessage" & strDBNoLock & ", " & strDbTable & "Author" & strDBNoLock & " " & _
	"WHERE " & strDbTable & "Author.Author_ID = " & strDbTable & "PMMessage.From_ID " & _
		"AND " & strDbTable & "PMMessage.PM_ID=" & lngPostID & "  " & _
		"AND " & strDbTable & "PMMessage.Author_ID=" & lngLoggedInUserID & ";"

	'Query the database
	rsCommon.Open strSQL, adoCon 
	
	
	'Read in the date of the reply pm
	dtmReplyPMDate = CDate(rsCommon("PM_Message_date"))
	
	'Make sure that the time and date format function isn't effected by the server time off set
	If strTimeOffSet = "-" Then
		dtmReplyPMDate = DateAdd("h", + intTimeOffSet, dtmReplyPMDate)
	ElseIf strTimeOffSet = "+" Then
		dtmReplyPMDate = DateAdd("h", - intTimeOffSet, dtmReplyPMDate)
	End If    
	
	
	'Build up the reply pm post
	strMessage = "<br /><br /><br />-- " & strTxtPreviousPrivateMessage & " --" & _
	"<br /><strong>" & strTxtSentBy & " :</strong> " & rsCommon("Username") & _
	"<br /><strong>" & strTxtSent & " :</strong> " & stdDateFormat(dtmReplyPMDate, True) & " " & strTxtAt & " " & TimeFormat(dtmReplyPMDate) & "<br /><br />"
	
	'Read in the quoted thread from the recordset
	strMessage = strMessage & rsCommon("PM_Message")
	
	'Clean up
	rsCommon.Close
End If





















'Make the post identical to before it was posted by removing border and target tags from the images and links
If NOT strMessage = "" Then 
	strMessage = Replace(strMessage, """ border=""0"" target=""_blank"">", """>", 1, -1, 1)
	strMessage = Replace(strMessage, """ border=""0"">", """>", 1, -1, 1)

	'If the message has been edited remove who edited the post
	If InStr(1, strMessage, "<edited>", 1) Then strMessage = removeEditorAuthor(strMessage)	
End If	





'If this is an edit or a quote check the user has permission
If strMode = "edit" OR strMode="editTopic" OR strMode = "quote" Then
	
	'Call the forum permissions function
	Call forumPermissions(intForumID, intGroupID)
	
	'If the user dosn't have permisison to view/edit/post/etc. then don't let them read the post
	If (strMode = "edit" OR strMode="editTopic") AND (blnAdmin = False AND blnModerator = False) Then
		
		If blnRead = False OR blnEdit = False Then strMessage = "Permission Denied!!"
	
	ElseIf strMode = "quote" Then
		
		If blnRead = False OR blnReply = False Then strMessage = "Permission Denied!!"
	
	End If
		
End If


	
'Reset Server Objects
Call closeDatabase()
%>
<html xmlns="http://www.w3.org/1999/xhtml" dir="<% = strTextDirection %>" lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=<% = strPageEncoding %>" />
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
<style>
html,body{border:0px;margin:1px;background-image:none;}
td {border:1px dotted #CCCCCC;}
</style>
</head>
<body class="WebWizRTEtextarea" leftmargin="1" topmargin="1" marginwidth="1" marginheight="1"><% 

'Display message
If NOT strMessage = "" Then 
	Response.Write(strMessage)

End If

 %></body>
</html>