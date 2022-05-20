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





'Declare variables
Dim sarryLatestTopics		'Holds the Last posts Feed recordset
Dim lngMessageID		'Holds the message ID
Dim dtmMessageDate		'Holds the message date
Dim lngAuthorID			'Holds the author ID
Dim strUsername			'Holds sthe authros user name
Dim strMessage			'Holds the post
Dim intLatestPostsMaxNo
Dim intRowColour	'Row colour number


'Set current record varible to 0
intCurrentRecord = 0
intRowColour = 0


'Set the max number of returned results
If blnMobileBrowser Then
	intLatestPostsMaxNo = 5
Else
	intLatestPostsMaxNo = 7
End If


'Get the last x posts from the database
strSQL = "" & _
"SELECT "
If strDatabaseType = "SQLServer" OR strDatabaseType = "Access" Then
	strSQL = strSQL & " TOP " & intLatestPostsMaxNo & " "
End If
strSQL = strSQL & _
"" & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Forum_name, " & strDbTable & "Topic.Topic_ID, " & strDbTable & "Topic.Subject, " & strDbTable & "Thread.Thread_ID, " & strDbTable & "Thread.Message_date, " & strDbTable & "Author.Author_ID, " & strDbTable & "Author.Username, " & strDbTable & "Thread.Message  " & _
"FROM " & strDbTable & "Forum, " & strDbTable & "Topic, " & strDbTable & "Author, " & strDbTable & "Thread " & _
"WHERE " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Topic.Forum_ID " & _
	"AND " & strDbTable & "Topic.Last_Thread_ID = " & strDbTable & "Thread.Thread_ID " & _
	"AND " & strDbTable & "Author.Author_ID = " & strDbTable & "Thread.Author_ID "
	
'Check permissions
strSQL = strSQL & _
	"AND (" & strDbTable & "Topic.Forum_ID " & _
		"IN (" & _
			"SELECT " & strDbTable & "Permissions.Forum_ID " & _
			"FROM " & strDbTable & "Permissions" & strDBNoLock & " " & _
			"WHERE (" & strDbTable & "Permissions.Group_ID = " & intGroupID & " OR " & strDbTable & "Permissions.Author_ID = " & lngLoggedInUserID & ") " & _
				"AND " & strDbTable & "Permissions.View_Forum = " & strDBTrue & _
		")" & _
	")"

'Don't include password protected forums
strSQL = strSQL & "AND (" & strDbTable & "Forum.Password = '' OR " & strDbTable & "Forum.Password Is Null) "

strSQL = strSQL & "AND (" & strDbTable & "Topic.Hide = " & strDBFalse & " AND " & strDbTable & "Thread.Hide = " & strDBFalse & ") " & _
"ORDER BY " & strDbTable & "Thread.Thread_ID DESC"

'mySQL limit operator
If strDatabaseType = "mySQL" Then
	strSQL = strSQL & " LIMIT " & intLatestPostsMaxNo
End If
strSQL = strSQL & ";"


'Set error trapping
On Error Resume Next

'Query the database
rsCommon.Open strSQL, adoCon


'If an error has occurred write an error to the page
If Err.Number <> 0 AND  strDatabaseType = "mySQL" Then
	Call errorMsg("An error has occurred while executing SQL query on database.<br />Please check that the MySQL Server version is 4.1 or above.", "get_latest_topic_data", "latest_topics_inc.asp")
ElseIf Err.Number <> 0 Then
	Call errorMsg("An error has occurred while executing SQL query on database.", "get_latest_topic_data", "latest_topics_inc.asp")
End If

'Disable error trapping
On Error goto 0



'Read in db results
If NOT rsCommon.EOF Then

	'Place the db results into an array
	sarryLatestTopics = rsCommon.GetRows()
	
End If

'Close the recordset
rsCommon.Close

'RS array lookup table
'0 = tblForum.Forum_ID
'1 = tblForum.Forum_name
'2 = tblTopic.Topic_ID
'3 = tblTopic.Subject
'4 = tblThread.Thread_ID
'5 = tblThread.Message_date
'6 = tblAuthor.Author_ID
'7 = tblAuthor.Username
'8 = tblThread.Message





'If there are records we need to display them
If isArray(sarryLatestTopics) Then
	
%>
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
 <tr class="tableLedger">
  <td colspan="3"><%
  	
  	If blnUrlRewrite Then
  		Response.Write("<a href=""new_forum_topics.html"">" & strTxtLatestForumPosts & "</a>")
	Else
		Response.Write("<a href=""active_topics.asp" & strQsSID1 & """>" & strTxtLatestForumPosts & "</a>")
	End If
  		
%>&nbsp;</td>
 </tr><%

	'If mobile view
	If blnMobileBrowser  = False Then

%>
 <tr class="tableSubLedger">
  <td width="32%"><% = strTxtTopics %></td>
  <td width="40%"><% = strTxtLastPost %></td>
  <td width="28%"><% = strTxtForum %></td>
 </tr><%
 
	End If	

	'Loop throug recordset to display the topics
	Do While intCurrentRecord <= Ubound(sarryLatestTopics, 2)

		'RS array lookup table
		'0 = tblForum.Forum_ID
		'1 = tblForum.Forum_name
		'2 = tblTopic.Topic_ID
		'3 = tblTopic.Subject
		'4 = tblThread.Thread_ID
		'5 = tblThread.Message_date
		'6 = tblAuthor.Author_ID
		'7 = tblAuthor.Username
		'8 = tblThread.Message


		'Read in db details for latest posts
		intForumID = CLng(sarryLatestTopics(0, intCurrentRecord))
		strForumName = sarryLatestTopics(1, intCurrentRecord)
		lngTopicID = CLng(sarryLatestTopics(2, intCurrentRecord))
		strSubject = sarryLatestTopics(3, intCurrentRecord)
		lngMessageID = CLng(sarryLatestTopics(4, intCurrentRecord))
		dtmMessageDate = CDate(sarryLatestTopics(5, intCurrentRecord))
		lngAuthorID = CLng(sarryLatestTopics(6, intCurrentRecord))
		strUsername = sarryLatestTopics(7, intCurrentRecord)
		strMessage = sarryLatestTopics(8, intCurrentRecord)

		'Srip the post to the first 150 chracters
		strMessage = removeHTML(strMessage, 150, true)
		
		'Set unread post count to 0
		intUnReadPostCount = 0
		
		'Count the number of unread posts in this forum
		If isArray(sarryUnReadPosts) AND dtmMessageDate > dtmLastVisitDate Then
			For intUnReadForumPostsLoop = 0 to UBound(sarryUnReadPosts,2)
				'Increament unread post count
				If CLng(sarryUnReadPosts(0,intUnReadForumPostsLoop)) = lngMessageID AND sarryUnReadPosts(3,intUnReadForumPostsLoop) = "1" Then intUnReadPostCount = intUnReadPostCount + 1
			Next
			
			'Get the text for unread post
			If intUnReadPostCount = 1 Then strNewPostText = strTxtNewPost Else strNewPostText = strTxtNewPosts	
		End If
		
		'Calculate the row colour
		intRowColour = intRowColour + 1
		
		
		'If mobile view
		If blnMobileBrowser Then
			
			If (intRowColour MOD 2 = 0 ) Then 
				Response.Write(vbCrLf & "  <tr class=""evenTableRow"">") 
			Else 
				Response.Write(vbCrLf & "  <tr class=""oddTableRow"">")
			End If
	   		 
			'Display the subject of the topic
			Response.Write("<td>")
							
			If blnUrlRewrite Then
				Response.Write("<a href=""" & SeoUrlTitle(strSubject, "") & "_topic" & lngTopicID & ".html" & strQsSID1 & "#" & lngMessageID & """>")
			Else
				Response.Write("<a href=""forum_posts.asp?TID=" & lngTopicID & "&PID=" & lngMessageID & strQsSID2 & SeoUrlTitle(strSubject, "&title=") & "#" & lngMessageID & """ title=""" & strMessage & """>")
			End If
			
			
			If blnBoldNewTopics AND intUnReadPostCount > 0 Then 'Unread topic subjects in bold
				Response.Write("<strong>" & strSubject & "</strong></a>")
				'Display the number of unread posts
				If intUnReadPostCount > 0 Then
			   		Response.Write(" <span class=""smText"">[" & strNewPostText & "]</span>") 
				End If
			Else
				Response.Write(strSubject & "</a>")
			End If
			
			'Display last post details
			Response.Write("<br /><span class=""smText"">" & strTxtBy & ": " & strUsername  & " " &  DateFormat(dtmMessageDate) & " " & strTxtAt & " " & TimeFormat(dtmMessageDate) & "</span>")
		
			
			
			Response.Write(vbCrLf & "  </tr>")
			
			
			
			
			
			

		'Not Mobile view
		Else
	
			If (intRowColour MOD 2 = 0 ) Then 
				Response.Write(vbCrLf & "  <tr class=""evenTableRow"">") 
			Else 
				Response.Write(vbCrLf & "  <tr class=""oddTableRow"">")
			End If
	   		 
			'Display the subject of the topic
			Response.Write("<td>")
							
			If blnUrlRewrite Then
				Response.Write("<a href=""" & SeoUrlTitle(strSubject, "") & "_topic" & lngTopicID & ".html" & strQsSID1 & "#" & lngMessageID & """ title=""" & strMessage & """>")
			Else
				Response.Write("<a href=""forum_posts.asp?TID=" & lngTopicID & "&PID=" & lngMessageID & strQsSID2 & SeoUrlTitle(strSubject, "&title=") & "#" & lngMessageID & """ title=""" & strMessage & """>")
			End If
			
			If blnBoldNewTopics AND intUnReadPostCount > 0 Then 'Unread topic subjects in bold
				Response.Write("<strong>" & strSubject & "</strong></a>")
				
			Else
				Response.Write(strSubject & "</a>")
			End If
			
			Response.Write("</td>")
			
			
				
			'Display who made the post and when
			Response.Write(vbCrLf & "   <td><span class=""smText"">" & _
			DateFormat(dtmMessageDate) & " " & strTxtAt & " " & TimeFormat(dtmMessageDate) & " " & _
			" " & strTxtBy & " <a href=""member_profile.asp?PF=" & lngAuthorID & strQsSID2 & """ class=""smLink"">" & strUsername & "</a> ")
			
			
			'If there are unread posts in the forum display differnt icon
			If intUnReadPostCount > 0 Then
				Response.Write("<a href=""forum_posts.asp?TID=" & lngTopicID & "&PID=" & lngMessageID & strQsSID2 & SeoUrlTitle(strSubject, "&title=") & "#" & lngMessageID & """><img src=""" & strImagePath & "view_unread_post." & strForumImageType & """ alt=""" & strTxtViewUnreadPost1 & """ title=""" & strTxtViewUnreadPost1 & """ /></a>")
									
			'Else there are no unread posts so display a normal last post link
			Else
				Response.Write("<a href=""forum_posts.asp?TID=" & lngTopicID & "&PID=" & lngMessageID & strQsSID2 & SeoUrlTitle(strSubject, "&title=") & "#" & lngMessageID & """><img src=""" & strImagePath & "view_last_post." & strForumImageType & """ alt=""" & strTxtViewPost & """ title=""" & strTxtViewPost & """ /></a>")
			End If
			
			Response.Write("</span>" & _
					"</td>")
	
	
	
			'Display the forum that was posted in
			Response.Write(vbCrLf & "   <td>")
			
			If blnUrlRewrite Then
				Response.Write("<a href=""" & SeoUrlTitle(strForumName, "") & "_forum" & intForumID & ".html" & strQsSID1 & """>" & strForumName & "</a>")
			Else
				Response.Write("<a href=""forum_topics.asp?FID=" & intForumID & strQsSID2 & SeoUrlTitle(strForumName, "&title=") & """>" & strForumName & "</a>")
			End If
			Response.Write(vbCrLf & "  </tr>")
		End If


  		'Increment the record position
  		intCurrentRecord = intCurrentRecord + 1

  	Loop
%>
</table>
<br /><%

End If

%>