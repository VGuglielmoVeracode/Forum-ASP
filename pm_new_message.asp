<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="language_files/pm_language_file_inc.asp" -->
<!--#include file="functions/functions_format_post.asp" -->
<!--#include file="includes/emoticons_inc.asp" -->
<!--#include file="functions/functions_send_mail.asp" -->
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






'Set the buffer to true
Response.Buffer = True

'Declare variables
Dim strToUsername	'Holds the username the pm message is sent to
Dim lngToUserID		'Holds author id of the person who the pm is for
Dim strSubject		'Holds the subject of the pm
Dim strMessage		'Holds the pm
Dim blnReadEmailNotify	'Holds if the user is to be notified when the message is read
Dim blnToUsernameOK	'Set to false if the to username is not found
Dim blnMaxPMsOK		'Set to false if the max number of private messages is exceeded
Dim blnMessageSent	'Set to true if the message is already sent
Dim strEmailSubject	'Holds the subject of the e-mail
Dim strEmailBody	'Holds the body of the e-mail message
Dim blnEmailSent	'set to true if an e-mail is sent
Dim blnBlocked		'Set to true if the user is blocked from messaging this person
Dim blnNoSubject	'Set to true if there is no subject to the PM
Dim strToEmail		'To email address
Dim blnPMNotify		'Set to true if the user wants notifying by emial
Dim intTmpPMno		'Temporary store for PM number the user has
Dim blnFloodControl	'Set to tru if flood control has been exceeded
Dim dtmFloodControlDate	'Holds the flood control date for the database search
Dim intSentPMs 		'Holds the number of PM sent
Dim lngPmMessageID	'Holds the new message ID
Dim lngPmMessageCount	'Holds the number of PM's received by the recepient
Dim objRegExp		'used for searches
Dim strBadWord		'Holds the bad words
Dim strBadWordReplace	'Holds the rplacment word for the bad word
Dim strSpamWord		'Holds the spam word
Dim blnSpamFilterFail	'Set to true if failed the spam filter
Dim intToGroupID



'Initilaise varaibles
blnSslEnabledPage = True
blnToUsernameOK = False
blnMaxPMsOK = True
blnMessageSent = False
blnBlocked = False
blnNoSubject = False
blnFloodControl = False
lngPmMessageCount = 0
blnSpamFilterFail = False



'If the user is user is using a banned IP redirect to an error page
If bannedIP() Then
	
	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp?M=IP" & strQsSID3)

End If


'If Priavte messages are not on then send them away
If blnPrivateMessages = False Then

	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("default.asp" & strQsSID1)
End If


'If the user is not allowed then send them away
If intGroupID = 2 OR blnActiveMember = False OR blnBanned Then
	
	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If

'Check the session ID to stop spammers using the email form
Call checkFormID(Request.Form("formID"))


'Read in the details for the pm
strSubject = Trim(Mid(Request.Form("subject"), 1, 41))
strMessage = Request.Form("Message")
blnReadEmailNotify = BoolC(Request.Form("email"))
strToUsername = Trim(Mid(Request.Form("member"), 1, 25))


'If the buddy text box is empty then read in the buddy from the list box
If strToUsername = "" Then strToUsername = Trim(Mid(Request.Form("selectMember"), 1, 15))

'Run the to username through the same SQL filer it was created under otherwise it may not match
strToUsername = formatSQLInput(strToUsername)

'Remove any tabs from subject
strSubject = Replace(strSubject, vbTab, " ")

'If there is no subject or message then don't post the message as won't be able to link to it
If strSubject = "" OR strMessage = "" Then blnNoSubject = True

'Get rid of scripting tags in the subject
strSubject = removeAllTags(strSubject)



'Check that the user the pm is being sent to exisits

'Initalise the SQL string with a query to read in the dteails of who the PM is to
strSQL = "SELECT " & strDbTable & "Author.Author_ID, " & strDbTable & "Author.Username, " & strDbTable & "Author.Group_ID, " & strDbTable & "Author.Author_email, " & strDbTable & "Author.PM_notify " & _
"FROM " & strDbTable & "Author" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Author.Username = '" & strToUsername & "';"

'Open the recordset
rsCommon.Open strSQL, adoCon

'If the to buddy is found in the database run the rest of the code
If NOT rsCommon.EOF Then

	'Username found so set to true
	blnToUsernameOK = True

	'Get details of who the PM is to
	lngToUserID = CLng(rsCommon("Author_ID"))
	strToEmail = rsCommon("Author_email")
	blnPMNotify = CBool(rsCommon("PM_notify"))
	intToGroupID = CBool(rsCommon("Group_ID"))


	'Don't let user send private message to guest account
	If (lngToUserID = 2 OR intToGroupID = 2) Then blnBlocked = True

	'Close the recordset
	rsCommon.Close



	'Check the user is not blocked from messaging this person

	'Initalise the SQL string with a query to see if the member has blocked this person from sending them a PM
	strSQL = "SELECT " & strDbTable & "BuddyList.Buddy_ID " & _
	"FROM " & strDbTable & "BuddyList" & strDBNoLock & " " & _
	"WHERE " & strDbTable & "BuddyList.Block = " & strDBTrue & " " & _
		"AND " & strDbTable & "BuddyList.Buddy_ID = " & lngLoggedInUserID & " " & _
		"AND " & strDbTable & "BuddyList.Author_ID = " & lngToUserID & ";"

	'Open the recordset
	rsCommon.Open strSQL, adoCon

	'If a record is returned then this user is blocked from messaging this person so don't send the pm, unless this is the forum admin
	If NOT rsCommon.EOF AND blnAdmin = False Then blnBlocked = True

	'Clean up
	rsCommon.Close




	'Check the user has not exceeded there allowed amount of private messages

	'Initalise the SQL string with a query to count the number of pm's the user has in their inbox
	strSQL = "SELECT Count(" & strDbTable & "PMMessage.PM_ID) AS CountOfPM " & _
	"FROM " & strDbTable & "PMMessage" & strDBNoLock & " " & _
	"WHERE " & strDbTable & "PMMessage.Inbox = " & strDBTrue & " " & _
		"AND " & strDbTable & "PMMessage.Author_ID = " & lngToUserID & ";"
	
	'Open the recordset
	rsCommon.Open strSQL, adoCon
	
	'If record returned we have the number of PM's this recepient has
	If NOT rsCommon.EOF Then
			
		'Get the number of PM's the recepient has received
		lngPmMessageCount = CLng(rsCommon("CountOfPM"))
	End If
		
	'Relese sever objects
	rsCommon.Close
	
	
		
	'If the number of PM's is more than max allowed in the inbox we need to either block or remove the PM from the inbox (built in admin accounts has no limits on the number of PM's they receive)
	If (lngPmMessageCount >= intPmInbox) AND intToGroupID > 1 Then 
			
		'If set to delete, then delete the oldest PM's the user has
		If  strPMoverAction = "delete" Then
			
			'Get PM's	
			strSQL = "SELECT " & strDbTable & "PMMessage.Inbox " & _
			"FROM " & strDbTable & "PMMessage" & strRowLock & " " & _
			"WHERE " & strDbTable & "PMMessage.Inbox = " & strDBTrue & " " & _
				"AND " & strDbTable & "PMMessage.Author_ID = " & lngToUserID & " " & _
			"ORDER BY " & strDbTable & "PMMessage.PM_ID ASC;"
			
			'Set the cursor	type property of the record set	to Forward Only
			rsCommon.CursorType = 0
		
			'Set the Lock Type for the records so that the record set is only locked when it is updated
			rsCommon.LockType = 3
			
			'Open the recordset
			rsCommon.Open strSQL, adoCon
			
			'Loop through and remove PM's from the inbox
			Do while NOT rsCommon.EOF
				
				'Update the database that this PM is no-longer available in the inbox
				rsCommon.Fields("Inbox") = strDBFalse
				rsCommon.Update
				
				'Remove 1 from the total count of PM's
				lngPmMessageCount = lngPmMessageCount - 1
				
				'If total count of PM's is the same or less than the max then exit loop
				'This is so we don't remove all the PM's from the inbox and only to the amount it is over by
				If lngPmMessageCount <= intPmInbox-1 Then Exit Do
				
				'Move to next record
				rsCommon.MoveNext
			Loop
			
			'Close rs
			rsCommon.Close
			
			
			'Delete all PM's that are not in the inbox or outbox for entire PM table to keep the PM table up-to-date
			'(May take a little longer, but removes all PM's that are no-longer being displayed in the system)	
			
			'SQL to delete all PM's that are not in the inbox or outbox
			strSQL = "DELETE FROM " & strDbTable & "PMMessage " & _
			"WHERE " & strDbTable & "PMMessage.Inbox = " & strDBFalse & " " & _
				"AND " & strDbTable & "PMMessage.Outbox = " & strDBFalse & ";"
					
			'Delete the message from the database
			adoCon.Execute(strSQL)
		
		'Else PM's are blocked
		Else
			blnMaxPMsOK = False
		
				
		End If
	End If


	
	
	
	
	'PM Flood control, make sure the user has not sent to many PM's
	If blnAdmin = False Then
		
		'Get the date with 1 hour taken off
		dtmFloodControlDate = internationalDateTime(DateAdd("h", -1, now()))
		
		'SQL Server doesn't like ISO dates with '-' in them, so remove the '-' part
		If strDatabaseType = "SQLServer" Then dtmFloodControlDate = Replace(dtmFloodControlDate, "-", "", 1, -1, 1)
	
		'Place the date in SQL safe # or '
		If strDatabaseType = "Access" Then
			dtmFloodControlDate = "#" & dtmFloodControlDate & "#"
		Else
			dtmFloodControlDate = "'" & dtmFloodControlDate & "'"
		End If
	
		'Initalise the SQL string with a query to read count the number of pm's the user has recieved
		strSQL = "SELECT Count(" & strDbTable & "PMMessage.PM_ID) AS CountOfSentPM " & _
		"FROM " & strDbTable & "PMMessage" & strDBNoLock & " " & _
		"WHERE " & strDbTable & "PMMessage.From_ID = " & lngLoggedInUserID & " " & _
			"AND " & strDbTable & "PMMessage.PM_Message_Date >= " & dtmFloodControlDate & ";"
	
		'Open the recordset
		rsCommon.Open strSQL, adoCon
	
		'If the user has exceeded the number of sent PM's in this hour don't let them send the PM
		If NOT rsCommon.EOF Then
			intSentPMs = CInt(rsCommon("CountOfSentPM"))
			
			If intSentPMs >= intPmFlood Then blnFloodControl = True
		End If

		'Relese sever objects
		rsCommon.Close
	End If
	
	
	
	
	
	'If the spam filters are enabled check the user is not sending spam
	If blnPmIgnoreSpamFilter = False  Then
		
		
		'Get the spam filters
		'Initalise the SQL string with a query to read in all the words from the smut table
		strSQL = "SELECT " & strDbTable & "Spam.* FROM " & strDbTable & "Spam" & strDBNoLock & ";"
	
		'Open the recordset
		rsCommon.Open strSQL, adoCon
		
		'Create regular experssions object
		Set objRegExp = New RegExp
	
		'Loop through all the spam filters
		Do While NOT rsCommon.EOF
			
			'Put the Spam into a string for imporoved perfoamnce
			strSpamWord = rsCommon("Spam")
			
			'Tell the regular experssions object what to look for
			With objRegExp
				.Pattern = strSpamWord
				.IgnoreCase = True
				.Global = True
			End With
			
			'Ignore errors, incase someone entered an incorrect bad word that breakes regular expressions
			On Error Resume Next
	
			'See if the spam is in the message
			If objRegExp.Execute(strMessage).Count > 0 Then blnSpamFilterFail = True
				
			
			 'Disable error trapping
			On Error goto 0
	
			'Move to the next word in the recordset
			rsCommon.MoveNext
		Loop
		
		'Distroy regular experssions object
		Set objRegExp = nothing
	
		'Release server objects
		rsCommon.Close
		
	End If
	
	
Else

	'Clean up
	rsCommon.Close
End If







'If the user to send to is found and they don't exceed max num of pm's (unless the sender is admin) then send the pm
If blnToUsernameOK AND blnMaxPMsOK AND blnBlocked = False AND blnNoSubject = False AND blnFloodControl = False AND blnSpamFilterFail = False Then
	

	'Place format posts posted with the WYSIWYG Editor
	If Request.Form("browser") = "RTE" Then

		'Call the function to format WYSIWYG posts
		strMessage = WYSIWYGFormatPost(strMessage)

	'Else standrd editor is used so convert forum codes
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




	'Replace swear words with other words with ***
	'Initalise the SQL string with a query to read in all the words from the smut table
	strSQL = "SELECT " & strDbTable & "Smut.* FROM " & strDbTable & "Smut" & strDBNoLock & ";"

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

		'Replace the swear words with the words in the database the swear words
		strMessage = objRegExp.Replace(strMessage, strBadWordReplace)
		strSubject = objRegExp.Replace(strSubject, strBadWordReplace)
		
		 'Disable error trapping
		On Error goto 0

		'Move to the next word in the recordset
		rsCommon.MoveNext
	Loop
	
	'Distroy regular experssions object
	Set objRegExp = nothing

	'Release server objects
	rsCommon.Close




	'Send (save) the private message
	'Initalise the SQL string (ADO is used for more security)
	strSQL = "SELECT" & strDBTop1 & " " & strDbTable & "PMMessage.* " & _
	"FROM " & strDbTable & "PMMessage" & strRowLock & " " & _
	"WHERE " & strDbTable & "PMMessage.Author_ID = " & lngToUserID & " " & _
	"ORDER BY " & strDbTable & "PMMessage.PM_ID DESC" & strDBLimit1 & ";"

	With rsCommon

		'Set the Lock Type for the records so that the record set is only locked when it is updated
		.LockType = 3

		'Open the recordset
		.Open strSQL, adoCon

		'Check to make sure the message is not already sent to the user
		If NOT .EOF Then
			If strMessage = rsCommon("PM_Message") Then blnMessageSent = True
		End IF

		'Save the pm
		If blnMessageSent = False Then

			'Add new record to recordset
			.AddNew
			
			.Fields("Author_ID") = lngToUserID
			.Fields("From_ID") = lngLoggedInUserID
			.Fields("Read_Post") = strDBFalse
			.Fields("PM_Tittle") = strSubject
			.Fields("PM_Message") = strMessage
			.Fields("PM_Message_date") = internationalDateTime(Now())
			'Check to see if they want e-mail notification of read pm
			If blnLoggedInUserEmail = True AND blnReadEmailNotify = True Then
				.Fields("Email_notify") = strDBTrue
			Else
				.Fields("Email_notify") = strDBFalse
			End If
			.Fields("Inbox") = strDBTrue
			.Fields("Outbox") = strDBTrue
			.Update
			
			'Update booleon that message is sent
			blnMessageSent = true
			
			'Requery to get message ID
			.Requery
			
			lngPmMessageID = .Fields("PM_ID")
			
		End If

		'Clean up
		.Close
	End With
	
	
	
	'When calling the updateUnreadPM function next, it can change the number of displayed un-read PM's for the PM sender
	'To prevent this bug, we store the number of un-read PM's for the sender in a temp store
	intTmpPMno = intNoOfPms
	
	'Update the number of unread PM's for the recepient
	Call updateUnreadPM(lngToUserID)
	
	'Restore the number of un-read PM's for the PM sender
	intNoOfPms = intTmpPMno



	'Update database with number of inbox PM's
	strSQL = "UPDATE " & strDbTable & "Author " & strRowLock & " " & _
	"SET " & strDbTable & "Author.Inbox_no_of_PM = " & lngPmMessageCount + 1 & " " & _
	"WHERE " & strDbTable & "Author.Author_ID = " & lngToUserID & ";"

	'Write the updated no. of PM's to the database
	adoCon.Execute(strSQL)




	'If the person has requested an email sent to them notifying them of the PM then send it
	If blnEmail AND blnPMNotify AND strToEmail <> "" Then

		'Set the subject
		strEmailSubject = strMainForumName & " " & strTxtNotificationPM

		'Initailise the e-mail body variable with the body of the e-mail
		strEmailBody = strTxtHi & " " & decodeString(strToUsername) & "," & _
		vbCrLf & vbCrLf & strTxtThisIsToNotifyYouThat & " '" & strLoggedInUsername & "' " & strTxtHasSentYouPM & ", '" & decodeString(strSubject) & "', " & strTxtOn & " " & strMainForumName & "." & _
		vbCrLf & vbCrLf & strTxtToViewThePrivateMessage & " " & strTxtForumClickOnTheLinkBelow & ": -" & _
                vbCrLf & vbCrLf & strForumPath & "pm_message.asp?ID=" & lngPmMessageID

		'Call the function to send the e-mail
		blnEmailSent = SendMail(strEmailBody, decodeString(strToUsername), decodeString(strToEmail), strWebsiteName, decodeString(strForumEmailAddress), decodeString(strEmailSubject), strMailComponent, false)
	End If
End If


'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & "<a href=""pm_welcome.asp" & strQsSID1 & """>" & strTxtPrivateMessenger & "</a>" & strNavSpacer & strTxtSendPrivateMessage


%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title><% = strTxtPrivateMessenger & " - " & strTxtSendPrivateMessage %></title>

<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write("<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Forums(TM) ver. " & strVersion & "" & _
vbCrLf & "Info: http://www.webwizforums.com" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->" & vbCrLf & vbCrLf)
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>

<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% = strTxtPrivateMessenger %></h1></td>
 </tr>
</table>
<br />
<table class="basicTable" cellspacing="0" cellpadding="0" align="center"> 
 <tr> 
  <td class="tabTable">
   <a href="pm_welcome.asp<% = strQsSID1 %>" title="<% = strTxtPrivateMessenger %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>messenger.<% = strForumImageType %>" border="0" alt="<% = strTxtPrivateMessenger %>" /> <% = strTxtMessenger %></a>
   <a href="pm_inbox.asp<% = strQsSID1 %>" title="<% = strTxtPrivateMessenger & " " & strTxtInbox %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>inbox_messages.<% = strForumImageType %>" border="0" alt="<% = strTxtPrivateMessenger & " " & strTxtInbox %>" /> <% = strTxtInbox %></a>
   <a href="pm_outbox.asp<% = strQsSID1 %>" title="<% = strTxtPrivateMessenger & " " & strTxtOutbox %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>sent_messages.<% = strForumImageType %>" border="0" alt="<% = strTxtPrivateMessenger & " " & strTxtOutbox %>" /> <% = strTxtOutbox %></a>
   <a href="pm_new_message_form.asp<% = strQsSID1 %>" title="<% = strTxtNewPrivateMessage %>" class="tabButtonActive">&nbsp;<img src="<% = strImagePath %>new_message.<% = strForumImageType %>" border="0" alt="<% = strTxtNewPrivateMessage %>" /> <% = strTxtNewMessage %></a>
  </td>
 </tr>
</table>
<br /><%

'If the pm has been sent show a message saying so
If blnMessageSent Then
	
%>
<table cellspacing="1" cellpadding="3" class="tableBorder" align="center">
    <tr class="tableLedger">
       <td align="left"><% = strTxtSendPrivateMessage %></td>
     </tr>
     <tr class="tableRow">
       <td><% =  strTxtYourPrivateMessage %> &quot;<% =  Server.HTMLEncode(strSubject) %>&quot;, <% = strTxtHasBeenSentTo & " " & strToUsername %>.
         <br /><br /><a href="pm_welcome.asp<% = strQsSID1 %>"><% =  strTxtReturnToYourPrivateMessenger %></a></td>
    </tr>
</table><%

'Else an error has occurred	
Else

%>
<form method="post" name="frmEditMessage" id="frmEditMessage" action="pm_new_message_form.asp?code=edit<% = strQsSID2 %>" onSubmit="return CheckForm();" onReset="return ResetForm();">
<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
  <tr>
    <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtError %></strong></td>
  </tr>
  <tr>
    <td><%
	'Display message to user
	If blnToUsernameOK = False Then
	
		'Display an error message
		Response.Write(strTxtYourPrivateMessage & " &quot;" & strSubject & "&quot;, " & strTxtHasNotBeenSent)
		Response.Write("<br /><br />" & strTxtTheUsernameCannotBeFound)
		Response.Write("<br /><br /><a href=""javascript:document.frmEditMessage.submit();"">" & strTxtAmendYourPrivateMessage & "</a>")
	
		'Save the pm details so they can be edited
		Response.Write(vbCrLf & "    <input type=""hidden"" name=""Subject"" id=""Subject"" value=""" & strSubject & """>")
		Response.Write(vbCrLf & "    <input type=""hidden"" name=""Buddy"" id=""Buddy"" value=""" & strToUsername & """>")
		Response.Write(vbCrLf & "    <input type=""hidden"" name=""PmMessage"" id=""PmMessage"" value=""" & Request.Form("Message") & """>")
	
	'If the message is blocked
	ElseIf blnBlocked Then
		'Display an error message
		Response.Write(strTxtYourPrivateMessage & " &quot;" & strSubject & "&quot;, " & strTxtHasNotBeenSent)
		Response.Write("<br /><br />" & strTxtYouAreBlockedFromSendingPMsTo & " " & strToUsername & ".")
		Response.Write("<br /><br /><a href=""pm_welcome.asp" & strQsSID1 & """>" & strTxtReturnToYourPrivateMessenger & "</a>")
	
	'Max PM's recieved
	ElseIf blnMaxPMsOK = False Then
		'Display an error message
		Response.Write(strTxtYourPrivateMessage & " &quot;" & strSubject & "&quot;, " & strTxtHasNotBeenSent)
		Response.Write("<br /><br />" & strToUsername & " " & strTxtHasExceededMaxNumPPMs & ".")
		Response.Write("<br /><br /><a href=""pm_welcome.asp" & strQsSID1 & """>" & strTxtReturnToYourPrivateMessenger & "</a>")
	
	'Flood Control
	ElseIf blnFloodControl Then
		'Display an error message
		Response.Write(strTxtYourPrivateMessage & " &quot;" & strSubject & "&quot;, " & strTxtHasNotBeenSent)
		Response.Write("<br /><br />" & strTxtYouAreOnlyPerToSend & " " & intPmFlood & " " & strTxtYouHaveExceededLimit & ".")
		Response.Write("<br /><br /><a href=""pm_welcome.asp" & strQsSID1 & """>" & strTxtReturnToYourPrivateMessenger & "</a>")
		
	'SPAM filter
	ElseIf blnSpamFilterFail Then
		'Display an error message
		Response.Write(strTxtYourPrivateMessage & " &quot;" & strSubject & "&quot;, " & strTxtHasNotBeenSent)
		Response.Write("<br /><br />" & strTxtYourMessageWasRejectedByTheSpamFilters & ".")
		Response.Write("<br /><br /><a href=""pm_welcome.asp" & strQsSID1 & """>" & strTxtReturnToYourPrivateMessenger & "</a>")
	
	
	'If there is no message body or subject display an error message
	ElseIf blnNoSubject Then
	
		'Display an error message
		Response.Write(strTxtYourPrivateMessage & " &quot;" & strSubject & "&quot;, " & strTxtHasNotBeenSent)
		Response.Write("<br /><br />" & strTxtYourMessageNoValidSubjectHeading)
		Response.Write("<br /><br /><a href=""pm_new_message_form.asp?code=edit" & strQsSID2 & """>" & strTxtAmendYourPrivateMessage & "</a>")
	
		'Save the pm details so they can be edited
		Response.Write(vbCrLf & "    <input type=""hidden"" name=""Subject"" id=""Subject"" value=""" & strSubject & """>")
		Response.Write(vbCrLf & "    <input type=""hidden"" name=""Buddy"" id=""Buddy"" value=""" & strToUsername & """>")
		Response.Write(vbCrLf & "    <input type=""hidden"" name=""PmMessage"" id=""PmMessage"" value=""" & Request.Form("Message") & """>")
	End If
%></td>
  </tr>
</table>
</form><%
 
End If

%>
<br />
<div align="center">
<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnLCode = True Then
	
	If blnTextLinks = True Then
		Response.Write("<span class=""text"" style=""font-size:10px"">Forum Software by <a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank""  style=""font-size:10px"">Web Wiz Forums&reg;</a> version " & strVersion & "</span>")
	Else
  		Response.Write("<a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank"" ><img src=""webwizforums_image.asp"" border=""0"" title=""Forum Software by Web Wiz Forums&reg; version " & strVersion& """ alt=""Forum Software by Web Wiz Forums&reg; version " & strVersion& """ /></a>")
	End If

	Response.Write("<br /><span class=""text"" style=""font-size:10px"">Copyright &copy;2001-2019 Web Wiz Ltd.</span>")
End If
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******

'Release server objects
Call closeDatabase()


'Display the process time
If blnShowProcessTime Then Response.Write "<span class=""smText""><br /><br />" & strTxtThisPageWasGeneratedIn & " " & FormatNumber(Timer() - dblStartTime, 3) & " " & strTxtSeconds & "</span>"
%>
 </div>
  <!-- #include file="includes/footer.asp" -->