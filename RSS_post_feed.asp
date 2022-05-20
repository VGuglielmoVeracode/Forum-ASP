<% @ Language=VBScript %>
<% Option Explicit %>
<!-- #include file="includes/global_variables_inc.asp" -->
<!-- #include file="includes/setup_options_inc.asp" -->
<!-- #include file="includes/version_inc.asp" -->
<!-- #include file="database/database_connection.asp" -->
<!-- #include file="functions/functions_common.asp" -->
<!-- #include file="functions/functions_filters.asp" -->
<!-- #include file="functions/functions_format_post.asp" -->
<!-- #include file="language_files/language_file_inc.asp" -->
<!-- #include file="language_files/RTE_language_file_inc.asp" -->
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

'Open database connection
Call openDatabase(strCon)

'Load in configuration data
Call getForumConfigurationData()


'If RSS is not enabled send the user away
If blnRSS = False Then

	'Clear server objects
	Call closeDatabase()

	'Redirect
	Response.Redirect("default.asp")
End If


'Include the date time format hear, incase a database hit is required if the date and time data is not in the web servers memory
%><!-- #include file="functions/functions_date_time_format.asp" --><%



'Declare variables
Dim sarryRssTopics	'Holds the RSS Feed recordset
Dim intCurrentRecord	'Holds the current record in the array
Dim strForumName	'Holds the forum name
Dim lngTopicID		'Holds the topic ID
Dim strSubject		'Holds the subject
Dim lngMessageID	'Holds the message ID
Dim dtmMessageDate	'Holds the message date
Dim lngAuthorID		'Holds the author ID
Dim strUsername		'Holds sthe authros user name
Dim strMessage		'Holds the post
Dim strStripedMessage	'Holds the first 30 chars of the message for a title
Dim strRssChannelTitle	'Holds the channel name
Dim strTimeZone		'Holds the time zone for the feed
Dim dtmLastEntryDate	'Holds the date of the last message
Dim intRSSLoopCounter	'Loop counter



'Set this to the time zone you require
strTimeZone = "+0000" 'See http://www.sendmail.org/rfc/0822.html#5 for list of time zones

'Initliase variables
lngTopicID = 0
strRssChannelTitle = strMainForumName


'Set the content type for feed
Response.ContentType = "application/xml"




'Read in the forum ID
If isNumeric(Request.QueryString("TID")) Then lngTopicID = LngC(Request.QueryString("TID"))



'If no topic ID boot the user
If lngTopicID = 0 Then

	'Clear server objects
	Call closeDatabase()

	'Redirect
	Response.Redirect("default.asp")
End If



'Get the last x posts from the database
strSQL = "" & _
"SELECT "
If strDatabaseType = "SQLServer" OR strDatabaseType = "Access" Then
	strSQL = strSQL & " TOP " & intRSSmaxResults & " "
End If
strSQL = strSQL & _
"" &  strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Forum_name, " & strDbTable & "Topic.Topic_ID, " & strDbTable & "Topic.Subject, " & strDbTable & "Thread.Thread_ID, " & strDbTable & "Thread.Message_date, " & strDbTable & "Author.Author_ID, " & strDbTable & "Author.Username, " & strDbTable & "Thread.Message  " & _
"FROM " & strDbTable & "Forum, " & strDbTable & "Topic, " & strDbTable & "Author, " & strDbTable & "Thread " & _
"WHERE " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Topic.Forum_ID " & _
	"AND " & strDbTable & "Topic.Topic_ID = " & strDbTable & "Thread.Topic_ID " & _
	"AND " & strDbTable & "Author.Author_ID = " & strDbTable & "Thread.Author_ID " & _
	"AND " & strDbTable & "Topic.Topic_ID = " & lngTopicID & " "

'Check permissions
strSQL = strSQL & _
	"AND (" & strDbTable & "Topic.Forum_ID " & _
		"IN (" & _
			"SELECT " & strDbTable & "Permissions.Forum_ID " & _
			"FROM " & strDbTable & "Permissions" & strDBNoLock & " " & _
			"WHERE (" & strDbTable & "Permissions.Group_ID = 2) " & _
				"AND " & strDbTable & "Permissions.View_Forum = " & strDBTrue & _
		")" & _
	")"

'Don't include password protected forums
strSQL = strSQL & "AND (" & strDbTable & "Forum.Password = '' OR " & strDbTable & "Forum.Password Is Null) "


strSQL = strSQL & "AND (" & strDbTable & "Topic.Hide = " & strDBFalse & " AND " & strDbTable & "Thread.Hide = " & strDBFalse & ") " & _
"ORDER BY " & strDbTable & "Thread.Message_date DESC"

'mySQL limit operator
If strDatabaseType = "mySQL" Then
	strSQL = strSQL & " LIMIT " & intRSSmaxResults
End If
strSQL = strSQL & ";"


'Set error trapping
On Error Resume Next

'Query the database
rsCommon.Open strSQL, adoCon


'If an error has occurred write an error to the page
If Err.Number <> 0 AND  strDatabaseType = "mySQL" Then
	Call errorMsg("An error has occurred while executing SQL query on database.<br />Please check that the MySQL Server version is 4.1 or above.", "get_RSS_post_data", "rss_post_feed.asp")
ElseIf Err.Number <> 0 Then
	Call errorMsg("An error has occurred while executing SQL query on database.", "get_RSS_post_data", "rss_post_feed.asp")
End If

'Disable error trapping
On Error goto 0



'Read in db results
If NOT rsCommon.EOF Then

	'Get the channel name
	strForumName = rsCommon("Forum_name")
	strSubject = rsCommon("Subject")
	dtmLastEntryDate = CDate(rsCommon("Message_date"))

	'Place the db results into an array
	sarryRssTopics = rsCommon.GetRows()

	'If the last post is more than 2 days ago change the time to live 6 hours
	If dtmLastEntryDate < DateAdd("d", -2, now()) Then
		intRssTimeToLive = 360
	'4 days Time to live 12 hours
	ElseIf dtmLastEntryDate < DateAdd("d", -4, now()) Then
		intRssTimeToLive = 720
	'1 week Time to live 24 hours
	ElseIf dtmLastEntryDate < DateAdd("w", -1, now()) Then
		intRssTimeToLive = 1440
	'4 weeks Time to live 2 days
	ElseIf dtmLastEntryDate < DateAdd("w", -4, now()) Then
		intRssTimeToLive = 2880
	End If
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



'Clear server objects
Call closeDatabase()


'Clean up channel name to prevent errors
strRssChannelTitle = Server.HTMLEncode(strRssChannelTitle) & " : " & strSubject
strForumName = Server.HTMLEncode(strForumName)
strMainForumName = Server.HTMLEncode(strMainForumName)



%><?xml version="1.0" encoding="<% = strPageEncoding %>" ?>
<?xml-stylesheet type="text/xsl" href="RSS_xslt_style.asp" version="1.0" ?>
<rss version="2.0" xmlns:WebWizForums="https://syndication.webwiz.net/rss_namespace/">
 <channel>
  <title><% = strRssChannelTitle %></title>
  <link><% = strForumPath %></link>
  <description><![CDATA[<% = strTxtThisIsAnXMLFeedOf %>; <% = strMainForumName & " : " & strForumName & " : " & strSubject %>]]></description><%

If blnLCode Then
	%>
  <copyright>Copyright (c) 2006-2013 Web Wiz Forums - All Rights Reserved.</copyright><%
End If

%>
  <pubDate><% = RssDateFormat(Now(), strTimeZone) %></pubDate>
  <lastBuildDate><% = RssDateFormat(dtmLastEntryDate, strTimeZone) %></lastBuildDate>
  <docs>http://blogs.law.harvard.edu/tech/rss</docs>
  <generator>Web Wiz Forums <% = strVersion %></generator>
  <ttl><% = intRssTimeToLive %></ttl>
  <WebWizForums:feedURL><% Response.Write(Replace(strForumPath, "http://", "")) %>RSS_post_feed.asp<% Response.Write("?TID=" & lngTopicID) %></WebWizForums:feedURL><%

'If there is a title image for the forum display it
If strTitleImage <> "" Then
	
	'If the title image is a http or https link then do not add the forum path to the image
	If InStr(strTitleImage, "http://") = 0 AND InStr(strTitleImage, "https://") = 0 Then strTitleImage = strForumPath & strTitleImage
%>
  <image>
   <title><![CDATA[<% = strMainForumName %>]]></title>
   <url><% = strTitleImage %></url>
   <link><% = strForumPath %></link>
  </image><%

End If



'If there are records we need to display them
If isArray(sarryRssTopics) Then

	'Loop throug recordset to display the topics
	Do While intCurrentRecord <= Ubound(sarryRssTopics, 2)

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


		'Read in db details for RSS feed
		intForumID = CLng(sarryRssTopics(0, intCurrentRecord))
		strForumName = sarryRssTopics(1, intCurrentRecord)
		lngTopicID = CLng(sarryRssTopics(2, intCurrentRecord))
		strSubject = sarryRssTopics(3, intCurrentRecord)
		lngMessageID = CLng(sarryRssTopics(4, intCurrentRecord))
		dtmMessageDate = CDate(sarryRssTopics(5, intCurrentRecord))
		lngAuthorID = CLng(sarryRssTopics(6, intCurrentRecord))
		strUsername = sarryRssTopics(7, intCurrentRecord)
		strMessage = sarryRssTopics(8, intCurrentRecord)


		'If the post contains a quote or code block then format it
		If InStr(1, strMessage, "[QUOTE=", 1) > 0 AND InStr(1, strMessage, "[/QUOTE]", 1) > 0 Then strMessage = formatUserQuote(strMessage)
		If InStr(1, strMessage, "[QUOTE]", 1) > 0 AND InStr(1, strMessage, "[/QUOTE]", 1) > 0 Then strMessage = formatQuote(strMessage)
		If InStr(1, strMessage, "[CODE]", 1) > 0 AND InStr(1, strMessage, "[/CODE]", 1) > 0 Then strMessage = formatCode(strMessage)
		If InStr(1, strMessage, "[HIDE]", 1) > 0 AND InStr(1, strMessage, "[/HIDE]", 1) > 0 Then strMessage = formatHide(strMessage)
		If InStr(1, strMessage, "[SPOILER]", 1) > 0 AND InStr(1, strMessage, "[/SPOILER]", 1) > 0 Then strMessage = formatSpoiler(strMessage, True)


		'If the post contains a flash link then format it
		If blnFlashFiles Then
			'Flash
			If InStr(1, strMessage, "[FLASH", 1) > 0 AND InStr(1, strMessage, "[/FLASH]", 1) > 0 Then strMessage = formatFlash(strMessage)
		End If
		
		'If YouTube
		If blnYouTube Then
			If InStr(1, strMessage, "[TUBE]", 1) > 0 AND InStr(1, strMessage, "[/TUBE]", 1) > 0 Then strMessage = formatYouTube(strMessage)
		End If


		'If the message has been edited parse the 'edited by' XML into HTML for the post
		If InStr(1, strMessage, "<edited>", 1) Then strMessage = editedXMLParser(strMessage)


		'Get the first 20 characters for a title for the message
		strStripedMessage = removeHTML(Trim(strMessage), 30, true)

		'Encode stripped message title to prevent XML parser error
		strStripedMessage = Server.HTMLEncode(strStripedMessage)



		'Format	the post to be sent with the e-mail
		strMessage = "<strong>" & strTxtAuthor & ":</strong> <a href=""" & strForumPath & "member_profile.asp?PF=" & lngAuthorID & """>" & strUsername & "</a>" & _
		"<br /><strong>" & strTxtSubject & ":</strong> " & sarryRssTopics(2, intCurrentRecord) & _
		"<br /><strong>" & strTxtPosted & "</strong> " & stdDateFormat(dtmMessageDate, True) & " " & strTxtAt & " " & TimeFormat(dtmMessageDate) & "<br /><br />" & _
		strMessage

		'Change	the path to the	emotion	symbols	to include the path to the images
		strMessage = Replace(strMessage, "src=""smileys/", "src=""" & strForumPath & "smileys/", 1, -1, 1)

		'Replace [] with HTML econded
		strMessage = Replace(strMessage, "[", "&#091;", 1, -1, 1)
		strMessage = Replace(strMessage, "]", "&#093;", 1, -1, 1)

		'Remove line breaks
		strMessage = Replace(strMessage, vbCrLf, "", 1, -1, 1)

%>
  <item>
   <title><![CDATA[<% = strSubject %> : <% = strStripedMessage %>]]></title>
   <link><% 
   		'If URL rewriting is enabled then build FURL link
   		If blnUrlRewrite Then
   			Response.Write(strForumPath & SeoUrlTitle(strSubject, "") & "_topic" & lngTopicID & "_post" & lngMessageID & ".html#" & lngMessageID)
   		Else
   			Response.Write(strForumPath & "forum_posts.asp?TID=" & lngTopicID & "&amp;PID=" & lngMessageID & SeoUrlTitle(strSubject, "&amp;title=") & "#" & lngMessageID)
   			
   		End If 
   %></link>
   <description>
    <![CDATA[<% = strMessage %>]]>
   </description>
   <pubDate><% = RssDateFormat(dtmMessageDate, strTimeZone) %></pubDate>
   <guid isPermaLink="true"><% 
   
   		'If URL rewriting is enabled then build FURL link
   		If blnUrlRewrite Then
   			Response.Write(strForumPath & SeoUrlTitle(strSubject, "") & "_topic" & lngTopicID & "_post" & lngMessageID & ".html#" & lngMessageID)
   		Else
   			Response.Write(strForumPath & "forum_posts.asp?TID=" & lngTopicID & "&amp;PID=" & lngMessageID & SeoUrlTitle(strSubject, "&amp;title=") & "#" & lngMessageID)
   			
   		End If 
   %></guid>
  </item> <%

  		'Increment the record position
  		intCurrentRecord = intCurrentRecord + 1

  	Loop

End If

%>
 </channel>
</rss>