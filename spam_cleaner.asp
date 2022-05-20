<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="functions/functions_date_time_format.asp" -->
<!--#include file="functions/functions_format_post.asp" -->
<!--#include file="functions/functions_edit_post.asp" -->
<!--#include file="includes/emoticons_inc.asp" -->
<!--#include file="language_files/admin_language_file_inc.asp" -->
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



'Set the timeout of the page
Server.ScriptTimeout = 1000


'Set the response buffer to true as we maybe redirecting
Response.Buffer = True


Dim saryRecordSet
Dim objXMLHTTP, objXmlDoc, objNode
Dim strDataStream
Dim lngUserProfileID
Dim blnSuspended
Dim strUsername
Dim lngPosts
Dim lngMemberPoints
Dim strAdminNotes
Dim strLastLoginIP
Dim strSignature
Dim strAvatar
Dim blnSuspendMember
Dim blnSubmitToStopForumSpam
Dim blnDeletePosts
Dim blnDeletePMs
Dim blnBlockIP
Dim strUserEmail
Dim blnStopForumSpamFound
Dim blnSpamUsername, blnSpamEmail, blnSpamIP
Dim blnIpInBlockList
Dim blnPMsSentOrReceived
Dim lngNumberOfPosts
Dim lngThreadID
Dim lngTopicID
Dim intForumIF
Dim lngLastPostID
Dim intCurrentRecord
Dim blnSfsSubmitted





'Initialise variables
blnSslEnabledPage = True
blnStopForumSpamFound = False
blnIpInBlockList = False
blnPMsSentOrReceived = False
lngNumberOfPosts = 0
intCurrentRecord = 0
blnSfsSubmitted = False


'Read in the member ID
lngUserProfileID = LngC(Request("PF"))


'If the person is not an admin or a moderator then send them away
If lngUserProfileID = "" OR bannedIP() OR  blnActiveMember = False OR blnBanned OR lngUserProfileID <= 2 Then
	
	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If



'First see if the user is a in a moderator group for any forum
'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT " & strDbTable & "Permissions.Moderate " & _
	"FROM " & strDbTable & "Permissions " & _
	"WHERE (" & strDbTable & "Permissions.Group_ID = " & intGroupID & " OR " & strDbTable & "Permissions.Author_ID = " & lngLoggedInUserID & ") AND  " & strDbTable & "Permissions.Moderate=" & strDBTrue & ";"
	

'Query the database
rsCommon.Open strSQL, adoCon

'If a record is returned then the user is a moderator in one of the forums
If NOT rsCommon.EOF Then blnModerator = True

'Clean up
rsCommon.Close


'If the user is not a moderator or admin then kick em
If blnAdmin = false AND  blnModerator = false Then
	
	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If


'If profile editing is disabled kick meoderator
  If blnAdmin = False AND blnModeratorProfileEdit = False Then
 	
 	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If




'Read the members details from the database
'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT " & strDbTable & "Author.Username, " & strDbTable & "Author.Author_email, " & strDbTable & "Author.No_of_posts, " & strDbTable & "Author.Points, " & strDbTable & "Author.Banned, " & strDbTable & "Author.Info, " & strDbTable & "Author.Login_IP, " & strDbTable & "Author.Signature, " & strDbTable & "Author.Avatar, " & strDbTable & "Author.No_of_PM, " & strDbTable & "Author.Homepage, " & strDbTable & "Author.Inbox_no_of_PM " & _
"FROM " & strDbTable & "Author " & _
"WHERE " & strDbTable & "Author.Author_ID = " & lngUserProfileID

'Set the cursor type property of the record set to Forward Only
rsCommon.CursorType = 0

'Set the Lock Type for the records so that the record set is only locked when it is updated
rsCommon.LockType = 3

'Query the database
rsCommon.Open strSQL, adoCon

'If not record boot
If rsCommon.EOF Then
	
	'Clean up
	rsCommon.Close
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
	
End If


'Read in the details of this member
strUsername = rsCommon("Username")
strUserEmail  = rsCommon("Author_email")
If isNull(rsCommon("No_of_posts")) Then lngPosts = 0 Else lngPosts = CLng(rsCommon("No_of_posts"))
If isNull(rsCommon("Points")) Then lngMemberPoints = 0 Else lngMemberPoints = CLng(rsCommon("Points"))
blnSuspended = CBool(rsCommon("Banned"))
strAdminNotes = rsCommon("Info")
strLastLoginIP = rsCommon("Login_IP")
strSignature = rsCommon("Signature")
strAvatar = rsCommon("Avatar")


'If admin note are blank prefill
If strAdminNotes = "" Then strAdminNotes = strAdminNotes & strLoggedInUsername & " " & strTxtCleanedSpammerSpamCleanerOn & " " & internationalDateTime(Now())


'Check to see if user is in the StopForumSpam Database
If blnStopForumSpam AND strStopForumSpamApiKey <> "" Then Call StopForumSpamLookup(strUserEmail, strUsername, strLastLoginIP)




'If this is a post back send the mail
If Request.Form("postBack") Then
	
	'Check the session ID to stop spammers using the email form
	Call checkFormID(Request.Form("formID"))
	
	'Read in form details
	blnSuspendMember = BoolC(Request.Form("banned"))
	blnSubmitToStopForumSpam = BoolC(Request.Form("sfs"))
	blnDeletePosts = BoolC(Request.Form("posts"))
	blnDeletePMs = BoolC(Request.Form("pms"))
	blnBlockIP = BoolC(Request.Form("ip"))
	
	strAvatar = Trim(Mid(Request.Form("txtAvatar"), 1, 95))
	strAdminNotes = Trim(Mid(removeAllTags(Request.Form("notes")), 1, 255))
	strSignature = Mid(Request.Form("signature"), 1, 210)
	
	
	'Call the function to format the signature
	strSignature = FormatPost(strSignature)
	strSignature = FormatForumCodes(strSignature)
	strSignature = HTMLsafe(strSignature)
	
	'Clean up avatar
	strAvatar = Trim(Mid(Request.Form("txtAvatar"), 1, 95))
	'If the avatar text box is empty then read in the avatar from the list box
	If strAvatar = "http://" OR strAvatar = "https://" OR strAvatar = "" Then strAvatar = Trim(Request.Form("SelectAvatar"))
	'If there is no new avatar selected then get the old one if there is one
	If strAvatar = "" Then strAvatar = Request.Form("oldAvatar")
	'If the avatar is the blank image then the user doesn't want one
	If strAvatar = strImagePath & "blank.gif" Then strAvatar = ""
		
		
	
	'******************************************
	'*** 	  Update tblAuthor Details	***
	'******************************************
	
	'Update member details first
	 With rsCommon
	 
	 	.Fields("Banned") = blnSuspendMember		
	 	If blnDeletePosts Then 
	 		.Fields("No_of_posts") = 0
	 		.Fields("Points") = 0
	 	End If	
		If blnDeletePMs Then 
			.Fields("No_of_PM") = 0
			.Fields("Inbox_no_of_PM") = 0
		End If
		.Fields("Avatar") = strAvatar
		.Fields("Homepage") = ""
		.Fields("Info") = strAdminNotes
		.Fields("Signature") = strSignature
		
		.Update
	 
	End With
	
	'Requery updated db and read in data again
	rsCommon.ReQuery
	
	'Read in the details of this member
	strUsername = rsCommon("Username")
	strUserEmail  = rsCommon("Author_email")
	If isNull(rsCommon("No_of_posts")) Then lngPosts = 0 Else lngPosts = CLng(rsCommon("No_of_posts"))
	If isNull(rsCommon("Points")) Then lngMemberPoints = 0 Else lngMemberPoints = CLng(rsCommon("Points"))
	blnSuspended = CBool(rsCommon("Banned"))
	strAdminNotes = rsCommon("Info")
	strLastLoginIP = rsCommon("Login_IP")
	strSignature = rsCommon("Signature")
	strAvatar = rsCommon("Avatar")
	
	
	 'If logging is enabled for moderators and banning member then log
	If blnLoggingEnabled AND blnModeratorLogging AND blnSuspendMember Then Call logAction(strLoggedInUsername, "SpamCleaner - Banned User " & strUsername)
	
	
	'Close the RS
	rsCommon.Close
	
	
	
	
	'******************************************
	'*** 	  	Block IP		***
	'******************************************
	
	'Block the members IP
	If blnBlockIP Then
		
		'Initalise the strSQL variable with an SQL statement to query the database to count the number of topics in the forums
		strSQL = "SELECT " & strDbTable & "BanList.Ban_ID, " & strDbTable & "BanList.IP, " & strDbTable & "BanList.Reason " & _
			"FROM " & strDbTable & "BanList" & strRowLock & " " & _
			"WHERE " & strDbTable & "BanList.IP = '" & strLastLoginIP & "';"
		
		'Set the cursor	type property of the record set	to Forward Only
		rsCommon.CursorType = 0
		
		'Set the Lock Type for the records so that the record set is only locked when it is updated
		rsCommon.LockType = 3
		
		'Query the database
		rsCommon.Open strSQL, adoCon
		
		'If no record returned then block IP
		If rsCommon.EOF Then
		
			'Update the recordset
			With rsCommon
			
				.AddNew
		
				'Update	the recorset
				.Fields("IP") = strLastLoginIP
				.Fields("Reason") = Trim(Mid(strUsername & " " & strTxtBlockedFromSpamFilter, 1, 40))
		
				'Update db
				.Update
				
				 'If logging is enabled for moderators record the IP blocking
			 	If blnLoggingEnabled AND blnModeratorLogging Then Call logAction(strLoggedInUsername, "SpamCleaner - IP Blocking: Username; " & strUsername  & " - IP; " & strLastLoginIP)
			End With
		End If
		
		'Close rs
		rsCommon.Close
	End If
	
	
	
	
	'******************************************
	'*** 	  	StopForumSpam		***
	'******************************************
	
	'If submit to StopForumSpam
	If blnSubmitToStopForumSpam AND blnStopForumSpam AND strStopForumSpamApiKey <> "" Then
			 	
		'If any part not found in the StopForumSpam Database submit the spammer
		If (blnSpamUsername = False OR blnSpamEmail = False OR blnSpamIP = False) AND strStopForumSpamApiKey <> "" Then
			 	
			'Submit spammer to StopForumSpam	
			Call StopForumSpamSubmit(strUserEmail, strUsername, strLastLoginIP)
			
			'Set submiited variable to true
			blnSfsSubmitted = True
			 	
			 	
			'If logging is enabled for moderators record the StopForumSpam Submissionj
			If blnLoggingEnabled AND blnModeratorLogging Then Call logAction(strLoggedInUsername, "SpamCleaner - StopForumSpam Submission: Username; " & strUsername  & " - Email; " & strUserEmail & " - IP; " & strLastLoginIP)
			 		
		End If
	 	
	   
	End If
	
	
	
	'******************************************
	'*** 	  	Delete PM's		***
	'******************************************
	
	'Delete PM's the meber has sent or received
	If blnDeletePMs Then
		
		'Two SQL delete queries used to prevent errors in mySQL
		
		'SQL to delete all PM's in user inbox
		strSQL = "DELETE FROM " & strDbTable & "PMMessage " & _
		"WHERE " & strDbTable & "PMMessage.Author_ID = " & lngUserProfileID & ";"
				
		'Delete the message from the database
		adoCon.Execute(strSQL)
		
		
		'SQL to delete PM's the user has sent
		strSQL = "DELETE FROM " & strDbTable & "PMMessage " & _
		"WHERE " & strDbTable & "PMMessage.From_ID = " & lngUserProfileID & ";"
				
		'Delete the message from the database
		adoCon.Execute(strSQL)
		
		 'If logging is enabled for moderators record the IP blocking
		If blnLoggingEnabled AND blnModeratorLogging Then Call logAction(strLoggedInUsername, "SpamCleaner - Deleted Private Messages sent and received for " & strUsername )
		
	End If
	
	
	
	
	
	
	
	'******************************************
	'*** 	  	Delete Posts		***
	'******************************************
	
	'Delete Posts this member has made
	If blnDeletePosts Then
		

		
		'Initalise the strSQL variable with an SQL statement to get the post from the database
		strSQL = "SELECT " & strDbTable & "Thread.Thread_ID, " & strDbTable & "Thread.Topic_ID " & _
		"FROM " & strDbTable & "Thread" & strDbNoLock & " " & _
		"WHERE " & strDbTable & "Thread.Author_ID = " & lngUserProfileID & ";"
		
		'Query the database
		rsCommon.Open strSQL, adoCon  
		
		'If NOT EOF
		If NOT rsCommon.EOF Then
			
			'Read the recordset into an array
			saryRecordSet = rsCommon.GetRows()
		
			'Close RS
			rsCommon.Close
		
		
			'Loop through posts and delete
			Do While intCurrentRecord <= Ubound(saryRecordSet,2)
				
				'Get the thread ID
				lngThreadID = CLng(saryRecordSet(0,intCurrentRecord))
				lngTopicID = CLng(saryRecordSet(1,intCurrentRecord))
	
				'Delete Post SQL
				strSQL = "DELETE FROM " & strDbTable & "Thread" & strRowLock & " " & _
				"WHERE " & strDbTable & "Thread.Thread_ID = "  & lngThreadID & ";"
				
				'Excute SQL
				adoCon.Execute(strSQL)
	
	
	
				'Check there are other Posts for the Topic, if not delete the topic as well	
				'Initalise the strSQL variable with an SQL statement to get the Threads from the database
				strSQL = "SELECT " & strDBTop1 & " " & strDbTable & "Thread.Thread_ID " & _
				"FROM " & strDbTable & "Thread" & strDBNoLock & " " & _
				"WHERE " & strDbTable & "Thread.Topic_ID = "  & lngTopicID & " " & _
				"ORDER BY " & strDbTable & "Thread.Message_date ASC" & strDBLimit1 & ";"
				
				'Query the database
				rsCommon.Open strSQL, adoCon
			
				
				'If there are posts left in the database for this topic get some details for them
				If NOT rsCommon.EOF Then
					
					'Get the post ID of the last post
					lngLastPostID = CLng(rsCommon("Thread_ID"))
				End If
				
				'Close the recordset
				rsCommon.Close
				
			
				
				'Read in details of the last topic to either update or delete (depends if any topics are left in db)
				
				'Initalise the strSQL variable with an SQL statement to get the topic from the database
				strSQL = "SELECT " & strDbTable & "Topic.* " & _
				"FROM " & strDbTable & "Topic" & strRowLock & " " & _
				"WHERE " & strDbTable & "Topic.Topic_ID = "  & lngTopicID & ";"
						
				'Set the cursor type property of the record set to Forward Only
				rsCommon.CursorType = 0
						
				'Set set the lock type of the recordset to optomistic while the record is deleted
				rsCommon.LockType = 3
						
				'Query the database
				rsCommon.Open strSQL, adoCon 
				
				
				'Get the forum ID for updating stats later
				intForumID = CInt(rsCommon("Forum_ID"))
				
				
				
				'If there are no other posts in the topic, delete the topic
				If lngLastPostID = 0 Then
					
					'If there is a poll and no more posts left delete the poll as well
					If CLng(rsCommon("Poll_ID")) <> 0 Then 
						
						'Delete the Poll choices 
						strSQL = "DELETE FROM " & strDbTable & "PollChoice " & strRowLock & " WHERE " & strDbTable & "PollChoice.Poll_ID = " & CLng(rsCommon("Poll_ID")) & ";" 
						
						'Write to database 
						adoCon.Execute(strSQL) 
						
						'Delete the Poll Votes 
						strSQL = "DELETE FROM " & strDbTable & "PollVote " & strRowLock & " WHERE " & strDbTable & "PollVote.Poll_ID = " & CLng(rsCommon("Poll_ID")) & ";" 
						
						'Write to database 
						adoCon.Execute(strSQL)
					
						'Delete the Poll 
						strSQL = "DELETE FROM " & strDbTable & "Poll " & strRowLock & " WHERE " & strDbTable & "Poll.Poll_ID = " & CLng(rsCommon("Poll_ID")) & ";" 
						
						'Write to database 
						adoCon.Execute(strSQL)  
					End If
					
					'delete any rating for this topic
					strSQL = "DELETE FROM " & strDbTable & "TopicRatingVote " & strRowLock & " " & _
					"WHERE " & strDbTable & "TopicRatingVote.Topic_ID = " & lngTopicID & ";"
					
					'Excute SQL
					adoCon.Execute(strSQL)
					
					
					'Delete Post Topic
					strSQL = "DELETE FROM " & strDbTable & "Topic" & strRowLock & " " & _
					"WHERE " & strDbTable & "Topic.Topic_ID = "  & lngTopicID & ";"
					
					'Excute SQL
					adoCon.Execute(strSQL)
					
					
					'Update the number of topics and posts in the database
					Call updateForumStats(intForumID)
					
					'Reset Server Objects
					rsCommon.Close
					
				
				'Else there are other posts in the topic, so let's update some details for the new last post
				Else 
					
					'Close Rs
					rsCommon.Close
					
					'Update the Topic Stats for this topic
					Call updateTopicStats(lngTopicID)
				End If
				
				'Update the number of topics and posts in the database
				Call updateForumStats(intForumID)
		
				
				'Move to next record
				intCurrentRecord = intCurrentRecord + 1
			Loop
		
			'If logging is enabled for moderators record the IP blocking
			If blnLoggingEnabled AND blnModeratorLogging Then Call logAction(strLoggedInUsername, "SpamCleaner - Deleted Posts for " & strUsername )
		
		Else
		
			'Close RS
			rsCommon.Close
		End If
		
	End If
	



'Else not a postback so close rs
Else
	
	rsCommon.Close
	

End If





'******************************************
'*** 	  Count number of Posts		***
'******************************************

'Initlise the sql statement
strSQL = "SELECT Count(" & strDbTable & "Thread.Thread_ID) AS CountOfPosts " & _
"FROM " & strDbTable & "Thread " & _
"WHERE " & strDbTable & "Thread.Author_ID = " & lngUserProfileID & ";"

'Query the database
rsCommon.Open strSQL, adoCon

'If records returned they has sent or recvived PM's
If NOT rsCommon.EOF Then lngNumberOfPosts = CLng(rsCommon("CountOfPosts"))


rsCommon.Close




'******************************************
'*** 	  Check for banned IP		***
'******************************************

'Check if members IP is listed
strSQL = "SELECT " & strDbTable & "BanList.Ban_ID, " & strDbTable & "BanList.IP, " & strDbTable & "BanList.Reason " & _
	"FROM " & strDbTable & "BanList" & strRowLock & " " & _
	"WHERE " & strDbTable & "BanList.IP = '" & strLastLoginIP & "';"
 
'Query the database
rsCommon.Open strSQL, adoCon
		
'If no record returned then block IP
If NOT rsCommon.EOF Then blnIpInBlockList = True
	
rsCommon.Close



'******************************************
'*** 	  Check number of PM's		***
'******************************************

'Initlise the sql statement
strSQL = "SELECT " & strDbTable & "PMMessage.PM_ID " & _
"FROM " & strDbTable & "PMMessage " & _
"WHERE " & strDbTable & "PMMessage.Author_ID = " & lngUserProfileID & " " & _
	"OR " & strDbTable & "PMMessage.From_ID = " & lngUserProfileID & ";"

'Query the database
rsCommon.Open strSQL, adoCon

'If records returned they has sent or recvived PM's
If NOT rsCommon.EOF Then blnPMsSentOrReceived = True
	
rsCommon.Close



'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtSpamCleaner

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title><% = strTxtSpamCleaner %></title>
<meta name="generator" content="Web Wiz Forums" />
<meta name="robots" content="noindex, follow" />

<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write("<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Forums(TM) ver. " & strVersion & "" & _
vbCrLf & "Info: http://www.webwizforums.com" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>
<script language="JavaScript">
function characterCounter(charNoBox, textFeild) {
	document.getElementById(charNoBox).value = document.getElementById(textFeild).value.length;
}
</script>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />		
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% = strTxtSpamCleaner %></h1></td>
 </tr>
</table>
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="right"><a href="member_profile.asp?PF=<% = lngUserProfileID & strQsSID2 %>"><% = strTxtProfile %></a> | <a href="member_control_panel.asp?PF=<% = lngUserProfileID %>&M=A<% = strQsSID2 %>"><% = strTxtEditMembersSettings %></a></td>
 </tr>
</table>
<form method="post" name="frmRegister" id="frmRegister" action="spam_cleaner.asp<% = strQsSID1 %>" onReset="return confirm('<% = strResetFormConfirm %>');">
 <table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
  <tr class="tableLedger">
   <td colspan="2"><% = strUsername %></td>
  </tr>
  <tr class="tableRow">
   <td width="50%"><br /><strong><% = strTxtUsername %></strong><br /><br /></td>
   <td width="50%"><br /><strong><% = strUsername %></stronmg><br /><br /></td>
   </tr>
   <tr class="tableRow">
    <td width="50%"><% = strTxtSuspendUser %><br /><span class="smText"><% = strTxtSuspendingMembersBetterThanDeleteing %></span><br /><br /></td>
    <td width="50%"><%
    	
    	'If members is already suspend display a message
    	If blnSuspended Then
    		
    		Response.Write("<img src=""" & strImagePath & "yes_green.png"" border=""0"" alt=""" & strTxtYes & """ /> " & strTxtThisMemberIsSuispended)
		Response.Write("<input type=""hidden"" name=""banned"" id=""banned"" value=""" & blnSuspended & """ />")
	
	'Else noit in database so let them be submitted
	Else
    		 %><% = strTxtYes %><input type="radio" name="banned" id="banned" value="True" checked="checked" tabindex="3" />&nbsp;&nbsp;<% = strTxtNo %><input type="radio" name="banned" id="banned" value="False" tabindex="4" /><%
    	 
    	 End If
    	
    	%></td>
   </tr><%
   
'If stopforumspam enabled
If blnStopForumSpam AND strStopForumSpamApiKey <> "" Then

%>   
   <tr class="tableRow">
    <td width="50%"><% = strTxtSubmitToStopForumSpam %><br /><span class="smText"><% = strTxtSubmitToStopForumSpamDatabaseToStopSpammer %>.</span><br /><br /></td>
    <td width="50%"><%
    	
    	'If already in StopForumSpamDatabase
    	If blnSpamUsername AND blnSpamEmail AND blnSpamIP Then
    		
    		Response.Write("<img src=""" & strImagePath & "yes_green.png"" border=""0"" alt=""" & strTxtYes & """ /> " & strTxtMembersDetailsFoundInStopForumSpamDatabase)
		Response.Write("<input type=""hidden"" name=""sfs"" id=""sfs"" value=""False"" />")
	
	'If submitted to stopforumspam let the user know	
	ElseIf blnSfsSubmitted Then
		
		Response.Write("<img src=""" & strImagePath & "yes_green.png"" border=""0"" alt=""" & strTxtYes & """ /> " & strTxtSubmittedToStopForumSpam)
		Response.Write("<input type=""hidden"" name=""sfs"" id=""sfs"" value=""False"" />")
	
	'Else not in database so let them be submitted
	Else
    		 %><% = strTxtYes %><input type="radio" name="sfs" id="sfs" value="True" checked="checked" tabindex="5" />&nbsp;&nbsp;<% = strTxtNo %><input type="radio" name="sfs" id="sfs" value="False" tabindex="6" /><%
    	 
    	 End If
    	 
    	 %></td>
   </tr><%
   
End If

%>   
   <tr class="tableRow">
    <td width="50%"><% = strTxtDeleteMembersPosts %><br /><br /></td>
    <td width="50%"><%
    	
    	'If member has no posts
    	If lngNumberOfPosts = 0 Then
    		
    		Response.Write("<img src=""" & strImagePath & "yes_green.png"" border=""0"" alt=""" & strTxtYes & """ /> " & strTxtNoPPostsCanBeFound)
		Response.Write("<input type=""hidden"" name=""posts"" id=""posts"" value=""False"" />")
	
	
	'Else noit in database so let them be submitted
	Else
    		 %><% = strTxtYes %><input type="radio" name="posts" id="posts" value="True" checked="checked" tabindex="7" />&nbsp;&nbsp;<% = strTxtNo %><input type="radio" name="posts" id="posts" value="False" tabindex="8" /><%
    	 
    	 End If
    	
    	%></td>
   </tr>
   
   <tr class="tableRow">
    <td width="50%"><% = strTxtDeleteMembersPrivateMessages %><br /><br /></td>
    <td width="50%"><%
    	
    	'If no PM's sent or received
    	If blnPMsSentOrReceived = False Then
    		
    		Response.Write("<img src=""" & strImagePath & "yes_green.png"" border=""0"" alt=""" & strTxtYes & """ /> " & strTxtNoPrivateMessagesCanBeFound)
		Response.Write("<input type=""hidden"" name=""pms"" id=""pms"" value=""False"" />")
	
	'Else noit in database so let them be submitted
	Else
    		 %><% = strTxtYes %><input type="radio" name="pms" id="pms" value="True" checked="checked" tabindex="9" />&nbsp;&nbsp;<% = strTxtNo %><input type="radio" name="pms" id="pms" value="False" tabindex="10" /><%
    	 
    	 End If
    	
    	%></td>
   </tr>
   
   <tr class="tableRow">
    <td width="50%"><% = strTxtBlockMembersIPpAddress %> <%
    	
    	If (blnAdmin OR (blnModerator AND blnModViewIpAddresses)) AND (strLastLoginIP <> "") Then Response.Write(" (" & strTxtIP & ": " & strLastLoginIP & " <a href=""https://network-tools.webwiz.net/ip-information.htm?ip=" & Server.URLEncode(strLastLoginIP) & """ target=""_blank""><img src=""" & strImagePath & "new_window.png"" alt=""" & strTxtIP & " " & strTxtInformation & """ title=""" & strTxtIP & " " & strTxtInformation & """ /></a>)") 	
    
    %>
    	
    	<br /><span class="smText"><% = strTxtBlockIPAddressFromRegPostInForum %>.</span><br /><br /></td>
    <td width="50%"><%
    	
    	'If already in IP Block List
    	If blnIpInBlockList Then
    		
    		Response.Write("<img src=""" & strImagePath & "yes_green.png"" border=""0"" alt=""" & strTxtYes & """ /> " & strTxtMembersLsstLoggedIpInBlockList)
		Response.Write("<input type=""hidden"" name=""ip"" id=""ip"" value=""False"" />")
	
	'Else noit in database so let them be submitted
	Else
    		 %><% = strTxtYes %><input type="radio" name="ip" id="ip" value="True" tabindex="11" />&nbsp;&nbsp;<% = strTxtNo %><input type="radio" name="ip" id="ip" value="False" tabindex="12" checked="checked" /><%
    	 
    	 End If
    	
    	%></td>
   </tr>
   
   
   
   <tr class="tableRow">
    <td valign="top"><% = strTxtSelectAvatar %><br /><span class="smText"><% = strTxtSelectAvatarDetails %>.</span></td>
    <td valign="top" height="2" >
    <table width="290" border="0" cellspacing="0" cellpadding="1">
     <tr>
      <td width="168">
       <select name="SelectAvatar" id="SelectAvatar" size="4" onchange="(avatar.src = SelectAvatar.options[SelectAvatar.selectedIndex].value) && (txtAvatar.value='http://') && (oldAvatar.value='')" tabindex="26">
        <option value="<% = strImagePath %>blank.gif"><% = strTxtNoneSelected %></option>
        <!-- #include file="includes/select_avatar.asp" -->
       </select>
      </td>
      <td width="122" align="center"><img src="<%

		'If there is an avatar then display it
		If strAvatar <> "" Then
		     	Response.Write(strAvatar)
		Else
			Response.Write(strImagePath & "blank.gif")
		End If
                %>" name="avatar" id="avatar" />
       <input type="hidden" name="oldAvatar" id="oldAvatar" value="<% = strAvatar %>"/></td>
      </tr>
      <tr>
       <td width="168"><input type="text" name="txtAvatar" id="txtAvatar" size="30" maxlength="95" value="<%

		'If the avatar is the persons own then display the link
		If InStr(1, strAvatar, "http://") > 0 OR InStr(1, strAvatar, "https://") > 0 Then
			Response.Write(strAvatar)
		Else
			Response.Write("http://")
		End If
        %>" onchange="oldAvatar.value=''" tabindex="27" /></td>
      <td width="122"><input type="button" name="preview" id="preview" value="<% = strTxtPreview %>" onclick="avatar.src=txtAvatar.value" tabindex="28" /><br /><br /></td>
     </tr>
    </table>
    </td>
   </tr>
   
   
   <tr class="tableRow">
    <td valign="top"><% = strTxtSignature %><br /><span class="smText"><% = strTxtSignatureLong %>&nbsp;(max 200 characters)
     <br />
     <br />
     <br />
     <a href="javascript:winOpener('BBcodes.asp<% = strQsSID1 %>','codes',1,1,610,500)" class="smLink"><% = strTxtForumCodes %></a> <% = strTxtForumCodesInSignature %></span>
    </td>
    <td valign="top" height="2">
     <textarea name="signature" id="signature" cols="30" rows="3" onKeyDown="characterCounter('sigChars', 'signature');" onKeyUp="characterCounter('sigChars', 'signature');" tabindex="29"><% = strSignature %></textarea>
     <br />
     <input size="3" value="0" name="sigChars" id="sigChars" maxlength="3" />
     <input onclick="characterCounter('sigChars', 'signature');" type="button" value="<% = strTxtCharacterCount %>" name="Count" />
     <br /><br />
    </td>
   </tr>
  
   <tr class="tableRow">
    <td valign="top"><% = strTxtAdminNotes %><br /><span class="smText"><% = strTxtAdminNotesAbout %>.</span></td>
    <td><textarea name="notes" id="notes" cols="30" rows="4" onKeyDown="characterCounter('notesChars', 'notes');" onKeyUp="characterCounter('notesChars', 'notes');" tabindex="55"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>><% = strAdminNotes %></textarea>
    	<br />
     <input size="3" value="0" name="notesChars" id="notesChars" maxlength="3" />
     <input onclick="characterCounter('notesChars', 'notes');" type="button" value="<% = strTxtCharacterCount %>" name="Count" />
    </td>
   </tr>
   <tr class="tableBottomRow">
    <td colspan="2" align="center">
     <input type="hidden" name="postback" id="postback" value="true" />
     <input type="hidden" name="PF" id="PF" value="<% = lngUserProfileID %>" />
     <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
     <input type="submit" name="Submit" id="Submit" value="<% Response.Write(strTxtSubmit) %>" onclick="return CheckForm();" tabindex="60" />
     <input type="reset" name="Reset" id="Reset" value="<% = strTxtResetForm %>" tabindex="61" />
    </td>
   </tr>
  </table>
 </form>
<br />


<%


'Clean up (done down here are session data may need to be saved)
Call closeDatabase()
%>
<br />
<div align="center"><%
   
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

'Display the process time
If blnShowProcessTime Then Response.Write "<span class=""smText""><br /><br />" & strTxtThisPageWasGeneratedIn & " " & FormatNumber(Timer() - dblStartTime, 3) & " " & strTxtSeconds & "</span>"
%></div>
<!-- #include file="includes/footer.asp" -->