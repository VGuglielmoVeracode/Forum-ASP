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




'Set the timeout of the page
Server.ScriptTimeout = 1000


'Set the response buffer to true as we maybe redirecting
Response.Buffer = True


'If in demo mode redirect
If blnDemoMode Then
	Call closeDatabase()
	Response.Redirect("admin_web_wiz_forums_premium.asp" & strQsSID1)
End If


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
Dim blnDeletePosts
Dim blnDeletePMs
Dim strUserEmail
Dim blnPMsSentOrReceived
Dim lngNumberOfPosts
Dim lngThreadID
Dim lngTopicID
Dim intForumIF
Dim lngLastPostID
Dim intCurrentRecord
Dim strReturn	'Holds the return page mode




'Initialise variables
blnSslEnabledPage = True
blnPMsSentOrReceived = False
lngNumberOfPosts = 0
intCurrentRecord = 0
strReturn = "UPD"



'Read in the member ID
lngUserProfileID = LngC(Request("PF"))






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

'Close RS
rsCommon.Close




'If this is a post back delete member (above member ID 2 so do not delete built in Admin and Guest accounts)
If Request.Form("postBack") AND lngUserProfileID > 2 Then
	
	
	'Read in form details
	blnDeletePosts = BoolC(Request.Form("posts"))
	blnDeletePMs = BoolC(Request.Form("pms"))
	
	
	
	
	'Delete the members buddy list
	'Initalise the strSQL variable with an SQL statement
	strSQL = "DELETE FROM " & strDbTable & "BuddyList WHERE (Author_ID = "  & lngUserProfileID & ") OR (Buddy_ID ="  & lngUserProfileID & ")"
		
	'Write to database
	adoCon.Execute(strSQL)	
		
		
	'Delete the members private msg's
	strSQL = "DELETE FROM " & strDbTable & "PMMessage WHERE (Author_ID ="  & lngUserProfileID & ")"
			
	'Write to database
	adoCon.Execute(strSQL)	
		
		
		
	'If option to delete sent PM's
	If blnDeletePMs Then
			
		'SQL to delete PM's the user has sent
		strSQL = "DELETE FROM " & strDbTable & "PMMessage WHERE " & strDbTable & "PMMessage.From_ID = " & lngUserProfileID & ";"
				
		'Delete the message from the database
		adoCon.Execute(strSQL)
			
		
	'Else option to keep sent PM's
	Else
		'Set all the users private messages to Guest account
		strSQL = "UPDATE " & strDbTable & "PMMessage SET From_ID = 2 WHERE (From_ID = "  & lngUserProfileID & ")"
				
		'Write to database
		adoCon.Execute(strSQL)
	End If
		
	
	'Remove users IP address from posts
	strSQL = "UPDATE " & strDbTable & "Thread SET IP_addr = '' WHERE (Author_ID = "  & lngUserProfileID & ")"
			
	'Write to database
	adoCon.Execute(strSQL)
	
		
	'Set all the users posts to the Guest account
	strSQL = "UPDATE " & strDbTable & "Thread SET Author_ID = 2 WHERE (Author_ID = "  & lngUserProfileID & ")"
			
	'Write to database
	adoCon.Execute(strSQL)
		
		
	'Set froums stats to the Guest account
	strSQL = "UPDATE " & strDbTable & "Forum SET Last_post_author_ID = 2 WHERE (Last_post_author_ID = "  & lngUserProfileID & ")"
			
	'Write to database
	adoCon.Execute(strSQL)
				
		
	'Delete the user from the email notify table
	strSQL = "DELETE FROM " & strDbTable & "EmailNotify WHERE (Author_ID = "  & lngUserProfileID & ")"
			
	'Write to database
	adoCon.Execute(strSQL)
		
		
	'Delete the user from forum permissions table
	strSQL = "DELETE FROM " & strDbTable & "Permissions WHERE (Author_ID = "  & lngUserProfileID & ")"
			
	'Write to database
	adoCon.Execute(strSQL)
		
		
		
	
	
		
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
					
					rsCommon.Close
				End If
				
				'Update the number of topics and posts in the database
				Call updateForumStats(intForumID)
		
				
				'Move to next record
				intCurrentRecord = intCurrentRecord + 1
			Loop
		
			
		
		Else
		
			'Close RS
			rsCommon.Close
		End If
		
	End If
	
	'Finally we can now delete the member from the forum
	strSQL = "DELETE FROM " & strDbTable & "Author WHERE (Author_ID = "  & lngUserProfileID & ")"
			
	'Write to database
	adoCon.Execute(strSQL)
		
	'Return page mode
	strReturn = "DEL"
	

End If


'If the delete process has not be completed
If strReturn <> "DEL" Then

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
End If

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Delete Forum Member</title>
<meta name="generator" content="Web Wiz Forums" />

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
<!-- #include file="includes/admin_header_inc.asp" -->
 <div align="center">
 <h1>Delete Forum Member</h1>
  <br /><br />
<% 
'The member has been deleted
If strReturn = "DEL" Then
	
	
	
%> 
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
  <tr class="tableLedger">
   <td colspan="2"><% = strUsername %></td>
  </tr>
  <tr class="tableRow">
   <td align="center"><br /><h2>The Member '<% = strUsername %> has been Deleted</h2><br /><br /></td>
   </tr>
   </td>
   </tr>
  </table>
<%
'Else display member details
Else
%> 
From here you can Delete a Members Data from your forum.
<br /><br />
<strong>Please note:</strong> If the Member has created allot of Posts, deletion can take some time to complete and may need to be run more than once if the Delete process times out.
<br /><br />
</div>
<form method="post" name="frmRegister" id="frmRegister" action="admin_delete_member.asp<% = strQsSID1 %>" onSubmit="return confirm('Are you sure you want to Permanently Delete this Member?');">
 <table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
  <tr class="tableLedger">
   <td colspan="2"><% = strUsername %></td>
  </tr>
  <tr class="tableRow">
   <td width="50%"><br /><strong>Username</strong><br /><br /></td>
   <td width="50%"><br /><strong><% = strUsername %></stronmg><br /><br /></td>
   </tr>
  
   <tr class="tableRow">
    <td width="50%">Delete Members Posts<br /><span class="smText">If you select NOT to delete this Members Posts their posts will be anonymised by removing the Username and IP Address of the Poster with all their posts being shown as being posted by 'Guest'.</span><br /><br /></td>
    <td width="50%"><%
    	
    	'If member has no posts
    	If lngNumberOfPosts = 0 Then
    		
    		Response.Write("<img src=""" & strImagePath & "yes_green.png"" border=""0"" alt=""" & strTxtYes & """ /> No Posts by " &  strUsername & " can be found")
		Response.Write("<input type=""hidden"" name=""posts"" id=""posts"" value=""False"" />")
	
	
	'Else not in database so let them be submitted
	Else
    		 %><% = strTxtYes %><input type="radio" name="posts" id="posts" value="True" tabindex="7" />&nbsp;&nbsp;<% = strTxtNo %><input type="radio" name="posts" id="posts" value="False" checked="checked" tabindex="8" /><%
    	 
    	 End If
    	
    	%></td>
   </tr>
   
   
    <tr class="tableRow">
    <td width="50%">Delete Sent Private Messsages<br /><span class="smText">Delete Private Messages that the Member has sent to others.</span><br /><br /></td>
    <td width="50%"><%
    	
    	'If no PM's sent or received
    	If blnPMsSentOrReceived = False Then
    		
    		Response.Write("<img src=""" & strImagePath & "yes_green.png"" border=""0"" alt=""" & strTxtYes & """ /> No Private Messages sent by " &  strUsername & " can be found")
		Response.Write("<input type=""hidden"" name=""pms"" id=""pms"" value=""False"" />")
	
	'Else noit in database so let them be submitted
	Else
    		 %><% = strTxtYes %><input type="radio" name="pms" id="pms" value="True" tabindex="9" />&nbsp;&nbsp;<% = strTxtNo %><input type="radio" name="pms" id="pms" value="False" checked="checked" tabindex="10" /><%
    	 
    	 End If
    	
    	%></td>
   </tr>
   

   <tr class="tableBottomRow">
    <td colspan="2" align="center">
     <input type="hidden" name="postback" id="postback" value="true" />
     <input type="hidden" name="PF" id="PF" value="<% = lngUserProfileID %>" />
     <input type="submit" name="Submit" id="Submit" value="Delete Member" onclick="return CheckForm();" tabindex="60" />
    </td>
   </tr>
  </table>
 </form>
<br />


<%
End If

'Clean up (done down here are session data may need to be saved)
Call closeDatabase()
%>

<!-- #include file="includes/admin_footer_inc.asp" -->
%>