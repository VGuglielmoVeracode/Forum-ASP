<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="functions/functions_date_time_format.asp" -->
<!--#include file="functions/functions_format_post.asp" -->
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

'Make sure this page is not cached
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"


'Dimension variables
Dim strForumName		'Holds the forum name
Dim lngNumberOfReplies		'Holds the number of replies for a topic
Dim lngTopicID			'Holds the topic ID
Dim strSubject			'Holds the topic subject
Dim strTopicIcon		'Holds the topic icon
Dim strTopicStartUsername 	'Holds the username of the user who started the topic
Dim lngTopicStartUserID		'Holds the users Id number for the user who started the topic
Dim lngNumberOfViews		'Holds the number of views a topic has had
Dim lngLastEntryMessageID	'Holds the message ID of the last entry
Dim strLastEntryUsername	'Holds the username of the last person to post a message in a topic
Dim lngLastEntryUserID		'Holds the user's ID number of the last person to post a meassge in a topic
Dim dtmLastEntryDate		'Holds the date the last person made a post in the topic
Dim intRecordPositionPageNum	'Holds the recorset page number to show the topics for
Dim intRecordLoopCounter	'Holds the loop counter numeber
Dim intTopicPageLoopCounter	'Holds the number of pages there are in the forum
Dim intLinkPageNum		'Holss the page number to link to
Dim intShowTopicsFrom		'Holds when to show the topics from
Dim strShowTopicsFrom		'Holds the display text from when the topics are shown from
Dim blnForumLocked		'Set to true if the forum is locked
Dim blnTopicLocked		'set to true if the topic is locked
Dim intPriority			'Holds the priority level of the topic
Dim intNumberOfTopicPages	'Holds the number of topic pages
Dim intTopicPagesLoopCounter	'Holds the number of loops
Dim blnHideTopic		'Holds if the topic is hidden
Dim strFirstPostMsg		'Holds the first posted message in the topic
Dim intForumReadRights		'Holds the read rights of the forum
Dim strForumPassword		'Holds the password for the forum
Dim strForumPaswordCode		'Holds the code for the password for the forum
Dim lngPollID			'Holds the topic poll id number
Dim dtmFirstEntryDate		'Holds the date of the first message
Dim strTableRowColour		'Holds the row colour for the table
Dim sarryTopics			'Holds the topics to display
Dim lngTotalRecords		'Holds the number of records in the topics array
Dim lngTotalRecordsPages	'Holds the total number of pages
Dim intStartPosition		'Holds the start poition for records to be shown
Dim intEndPosition		'Holds the end poition for records to be shown
Dim intCurrentRecord		'Holds the current record position
Dim strSearchKeywords		'Holds the search keywods for highlighting
Dim strSearchID			'Holds the search ID
Dim sarySearchWord		'Holds the search words
Dim sarySearchIndex		'Holds the details of the search array
Dim strSearchMemID		'Holds the ID of the memebr who ran the search
Dim dtmSearchDateCreated	'Holds the date the search was created
Dim dblSearchProcessTime	'Holds the time taken to process the search
Dim intPageLinkLoopCounter	'Holds the loop counter for mutiple page links
Dim dtmEventDate		'Holds the date if this is a calendar event
Dim dtmEventDateEnd		'Holds the date if this is a calendar event
Dim intUnReadPostCount		'Holds the count for the number of unread posts in the forum
Dim intUnReadForumPostsLoop	'Loop to count the number of unread posts in a forum
Dim intMovedForumID
Dim dblTopicRating		'Holds the rating for a topic
Dim lngTopicVotes		'Number of votes a topic receives
Dim strNewPostText
Dim strSeoSubject		'Holds the subject
Dim dblGetLastPostID


'Initlaise variables
dblGetLastPostID = CDbl(Now())




'If we have not yet checked for unread posts since last visit run it now
If Session("dtmUnReadPostCheck") = "" Then 
	Call UnreadPosts()
'Read in array if at application level
ElseIf isArray(Application("sarryUnReadPosts" & strSessionID)) Then  
	sarryUnReadPosts = Application("sarryUnReadPosts" & strSessionID)
'Read in the unread posts array	
ElseIf isArray(Session("sarryUnReadPosts")) Then 
	sarryUnReadPosts = Session("sarryUnReadPosts")
End If



'If this is the first time the page is displayed then the Forum Topic record position is set to page 1
If isNumeric(Request.QueryString("PN")) = false Then
	intRecordPositionPageNum = 1
ElseIf Request.QueryString("PN") < 1 Then
	intRecordPositionPageNum = 1

'Else the page has been displayed before so the Forum Topic record postion is set to the Record Position number
Else
	intRecordPositionPageNum = IntC(Request.QueryString("PN"))
End If


'Read in the querystrings
strSearchID = Request.QueryString("SearchID")
strSearchKeywords = Request.QueryString("KW")

'Split up the keywords to be searched
sarySearchWord = Split(Trim(strSearchKeywords), " ")



'Read in the search index array to see if user has permission to view this search and update the time to live date
'Place the application level search results array into a temporary dynaimic array
sarySearchIndex = Application(strAppPrefix & "sarySearchIndex")

'Array dimension lookup table
' 0 = Search ID
' 1 = IP
' 2 = Date/time last run
' 3 = Member ID
' 4 = Date/time search created
' 5 = time taken to run search

'Iterate through the array to find the index data for this search
If isArray(sarySearchIndex) Then
	For intCurrentRecord = 0 To UBound(sarySearchIndex, 2)
		
		'Find the array data for this search
		If sarySearchIndex(0, intCurrentRecord) = strSearchID Then
			'Read the the data we are using from the index
			strSearchMemID = CLng(sarySearchIndex(3, intCurrentRecord))
			dtmSearchDateCreated = CDate(sarySearchIndex(4, intCurrentRecord))
			dblSearchProcessTime = CDbl(sarySearchIndex(5, intCurrentRecord))
			
			'Exit loop as we have the data we need
			Exit For
		End If
	Next 
End If


''If the user is the person who created the search they have permission to view it
'Read in the search array and update the time to live
If strSearchMemID = lngLoggedInUserID Then
	
	'Read in the search array
	sarryTopics = Application(strSearchID)
	
	'Update the time to live
	sarySearchIndex(2, intCurrentRecord) = internationalDateTime(Now())
	
	'Update search index application array
	'Lock the application so that no other user can try and update the application level variable at the same time
	Application.Lock
				
	'Update the application level variables
	Application(strAppPrefix & "sarySearchIndex") = sarySearchIndex
				
	'Unlock the application
	Application.UnLock
End If

'Reset current record variable
intCurrentRecord = 0




'Array Look Up table
'0 = tblForum.Forum_ID
'1 = tblForum.Forum_name
'2 = tblForum.Password
'3 = tblForum.Forum_code
'4 = tblTopic.Topic_ID
'5 = tblTopic.Poll_ID
'6 = tblTopic.Moved_ID
'7 = tblTopic.Subject
'8 = tblTopic.Icon
'9 = tblTopic.Start_Thread_ID
'10 = tblTopic.Last_Thread_ID
'11 = tblTopic.No_of_replies
'12 = tblTopic.No_of_views
'13 = tblTopic.Locked
'14 = tblTopic.Priority
'15 = tblTopic.Hide
'16 = tblThread.Message_date
'17 = tblThread.Message, 
'18 = tblThread.Author_ID, 
'19 = tblAuthor.Username, 
'20 = LastThread.Message_date, 
'21 = LastThread.Author_ID, 
'22 = LastAuthor.Username
'23 = tblTopic.Event_date
'24 = tblTopic.Event_date_end
'25 = tblTopic.Rating
'26 = tblTopic.Rating_Votes


'Read in some details of the topics
If isArray(sarryTopics) Then 
	
	'Count the number of records
	lngTotalRecords = Ubound(sarryTopics,2) + 1

	'Count the number of pages for the topics using '\' so that any fraction is omitted 
	lngTotalRecordsPages = lngTotalRecords \ intTopicPerPage
	
	'If there is a remainder or the result is 0 then add 1 to the total num of pages
	If lngTotalRecords Mod intTopicPerPage > 0 OR lngTotalRecordsPages = 0 Then lngTotalRecordsPages = lngTotalRecordsPages + 1
		
	'Start position
	intStartPosition = ((intRecordPositionPageNum - 1) * intTopicPerPage)

	'End Position
	intEndPosition = intStartPosition + intTopicPerPage
		
	'Get the start position
	intCurrentRecord = intStartPosition
End If



'Page to link to for mutiple page (with querystrings if required)
strLinkPage = "search_results_topics.asp?SearchID=" & Server.URLEncode(strSearchID) & "&KW=" & Server.URLEncode(strSearchKeywords) & "&"


'If active users is enabled update the active users application array
If blnActiveUsers Then
	'Call active users function
	saryActiveUsers = activeUsers(strTxtSearchingForums, strTxtSearchingFor & ": &#8216;" & Server.HTMLEncode(strSearchKeywords) & "&#8217;", "search_form.asp?KW=" & Server.URLEncode(strSearchKeywords), 0)
End If


'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & "<a href=""search_results_topics.asp?SearchID=" & Server.URLEncode(strSearchID) & "&KW=" & Server.URLEncode(strSearchKeywords) & strQsSID2 & """>" & strTxtSearchResults & "</a>"

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title><% = strMainForumName & " - " & strTxtSearchResults & " - " & Server.HTMLEncode(strSearchKeywords) %></title>

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

<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% = strTxtSearchResults %></h1>
  <br /><%

'Display some text on search
If lngTotalRecords > 0 Then
	
	Response.Write(strTxtSearchResults & " ")
	
	'If this is a keyword search display keywrds
	If strSearchKeywords <> "" Then 
		Response.Write(strTxtFor & " '" & Server.HTMLEncode(strSearchKeywords) & "' ")
	End If
	
	Response.Write(strTxtHasFound & " " & lngTotalRecords & " " & strTxtResultsIn & " " & dblSearchProcessTime & " " & strTxtSecounds & ".")
	
	Response.Write("<br /><span class=""smText"">" & strTxtThisSearchWasProcessed & " " & DateFormat(dtmSearchDateCreated) & " " & strTxtAt & " " & TimeFormat(dtmSearchDateCreated) & ".</span>")
End If
%></td>
   <td align="right" valign="top" nowrap><!-- #include file="includes/page_link_inc.asp" --></td>
 </tr>
</table>
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center"><%

'Display column headings if not a mobile
If blnMobileBrowser = False Then

%>
 <tr class="tableLedger">
  <td width="5%">&nbsp;</td>
  <td width="50%"><div style="float:left;"><% = strTxtTopics & " / " & strTxtThreadStarter %></div></div><% 
   	'If rating is enabled
   	If blnTopicRating Then 
   		%><div style="float:right;"><% = strTxtRating %>&nbsp;</div><%
   		
   	End If 
   %></td>
  <td width="10%" align="center"><% = strTxtReplies %></td>
  <td width="10%" align="center"><% = strTxtViews %></td>
  <td width="30%"><% = strTxtLastPost %></td>
 </tr><%
End If



'If there are no search results display an error msg
If lngTotalRecords <= 0 Then
	
	'If there are no search results to display then display the appropriate error message
	Response.Write vbCrLf & " <tr class=""tableRow""><td colspan=""6"" align=""center""><br />" & strTxtSearhExpiredOrNoPermission & " <a href=""search_form.asp?KW=" & Server.URLEncode(strSearchKeywords) & strQsSID2 & """>" & strTxtCreateNewSearch & "</a><br /><br /></td></tr>"




'Disply any search results in the forum
Else


	'Do....While Loop to loop through the recorset to display the forum topics
	Do While intCurrentRecord < intEndPosition

		'If there are no topic records left to display then exit loop
		If intCurrentRecord >= lngTotalRecords Then Exit Do
			
			
		'SQL Query Array Look Up table
		'0 = tblForum.Forum_ID
		'1 = tblForum.Forum_name
		'2 = tblForum.Password
		'3 = tblForum.Forum_code
		'4 = tblTopic.Topic_ID
		'5 = tblTopic.Poll_ID
		'6 = tblTopic.Moved_ID
		'7 = tblTopic.Subject
		'8 = tblTopic.Icon
		'9 = tblTopic.Start_Thread_ID
		'10 = tblTopic.Last_Thread_ID
		'11 = tblTopic.No_of_replies
		'12 = tblTopic.No_of_views
		'13 = tblTopic.Locked
		'14 = tblTopic.Priority
		'15 = tblTopic.Hide
		'16 = tblThread.Message_date
		'17 = tblThread.Message, 
		'18 = tblThread.Author_ID, 
		'19 = tblAuthor.Username, 
		'20 = LastThread.Message_date, 
		'21 = LastThread.Author_ID, 
		'22 = LastAuthor.Username
		'23 = tblTopic.Event_date
		'24 = tblTopic.Event_date_end
		'25 = tblTopic.Rating
		'26 = tblTopic.Rating_Votes


		'Read in Topic details from the database
		intForumID = CInt(sarryTopics(0,intCurrentRecord))
		strForumPassword = sarryTopics(2,intCurrentRecord)
		strForumPaswordCode = sarryTopics(3,intCurrentRecord)
		
		
		
		'Read in Topic details from the database
		lngTopicID = CLng(sarryTopics(4,intCurrentRecord))
		lngPollID = CLng(sarryTopics(5,intCurrentRecord))
		strSubject = sarryTopics(7,intCurrentRecord)
		strTopicIcon = sarryTopics(8,intCurrentRecord)
		lngNumberOfReplies = CLng(sarryTopics(11,intCurrentRecord))
		lngNumberOfViews = CLng(sarryTopics(12,intCurrentRecord))
		blnTopicLocked = CBool(sarryTopics(13,intCurrentRecord))
		intPriority = CInt(sarryTopics(14,intCurrentRecord))
		blnHideTopic = CBool(sarryTopics(15,intCurrentRecord))
		dtmEventDate = sarryTopics(23,intCurrentRecord)
		dtmEventDateEnd = sarryTopics(24,intCurrentRecord)
		If isNumeric(sarryTopics(25,intCurrentRecord)) Then dblTopicRating = CDbl(sarryTopics(25,intCurrentRecord)) Else dblTopicRating = 0
		If isNumeric(sarryTopics(26,intCurrentRecord)) Then lngTopicVotes = CLng(sarryTopics(26,intCurrentRecord)) Else lngTopicVotes = 0
		
		'Read in the first post details
		dtmFirstEntryDate = CDate(sarryTopics(16,intCurrentRecord))
		strFirstPostMsg = Mid(sarryTopics(17,intCurrentRecord), 1, 275)
		lngTopicStartUserID = CLng(sarryTopics(18,intCurrentRecord))
		strTopicStartUsername = sarryTopics(19,intCurrentRecord)
	
		'Read in the last post details
		lngLastEntryMessageID = CLng(sarryTopics(10,intCurrentRecord))
		dtmLastEntryDate = CDate(sarryTopics(20,intCurrentRecord))
		lngLastEntryUserID = CLng(sarryTopics(21,intCurrentRecord))
		strLastEntryUsername = sarryTopics(22,intCurrentRecord)
		
		
		'Get the SEO Title
		strSeoSubject = strSubject 
		
		'Highlight the search words
		strSubject = searchHighlighter(strSubject, sarySearchWord)
		
		'Clean up input to prevent XXS hack
		strSubject = formatInput(strSubject)
		
		
		
		'If forum is passworded and not logged into forum display that password is required
		If strForumPassword <> "" and getCookie("fID", "Forum" & intForumID) <> strForumPaswordCode Then 
			strFirstPostMsg = strTxtPasswordRequiredViewPost
			strSubject = strTxtPasswordRequiredViewPost
			strTopicStartUsername = strTxtNotGiven
			lngTopicStartUserID = 2
			strLastEntryUsername = strTxtNotGiven
			lngLastEntryUserID = 2
		End If
		

		'If the forum name is different to the one from the last forum display the forum name
		If sarryTopics(1,intCurrentRecord) <> strForumName Then

			'Give the forum name the new forum name
			strForumName = sarryTopics(1,intCurrentRecord)

			'Display the new forum name
			Response.Write vbCrLf & " <tr class=""tableSubLedger""><td colspan=""7""><a href=""forum_topics.asp?FID=" & intForumID & strQsSID2 & """>" & strForumName & "</a></td></tr>"
		End If
		
		
		
		'Remove HTML from message for subject link title
		strFirstPostMsg = removeHTML(strFirstPostMsg, 150, true)
		
		
		'Clean up input to prevent XXS hack
		strFirstPostMsg = formatInput(strFirstPostMsg)


		'Unread Posts *********
		intUnReadPostCount = 0
					
		'If there is a newer post than the last time the unread posts array was initilised run it again
		If dtmLastEntryDate > CDate(Session("dtmUnReadPostCheck")) Then Call UnreadPosts()
						
		'Count the number of unread posts in this forum
		If isArray(sarryUnReadPosts) AND dtmLastEntryDate > dtmLastVisitDate Then
			For intUnReadForumPostsLoop = 0 to UBound(sarryUnReadPosts,2)
				'Increament unread post count
				If CLng(sarryUnReadPosts(1,intUnReadForumPostsLoop)) = lngTopicID AND sarryUnReadPosts(3,intUnReadForumPostsLoop) = "1" Then intUnReadPostCount = intUnReadPostCount + 1
			Next
			
			'Get the text for unread post
			If intUnReadPostCount = 1 Then strNewPostText = strTxtNewPost Else strNewPostText = strTxtNewPosts	
		End If
	
		
		
		'Calculate the row colour
		If intCurrentRecord MOD 2=0 Then strTableRowColour = "evenTableRow" Else strTableRowColour = "oddTableRow"
		
		'If this is a hidden post then change the row colour to highlight it
		If blnHideTopic Then strTableRowColour = "hiddenTableRow"



		'If mobile browser display different content
		If blnMobileBrowser Then
			
			'Write the HTML of the Topic descriptions as hyperlinks to the Topic details and message
			Response.Write(vbCrLf & "  <tr class=""" & strTableRowColour & """>")
			
			'Display the subject of the topic
			Response.Write("<td><a href=""forum_posts.asp?TID=" & lngTopicID)
			If intPriority = 3 Then Response.Write("&amp;FID=" & intForumID & "&amp;PR=3")
			Response.Write("" & strQsSID2 & SeoUrlTitle(strSubject, "&title=") & """ title=""" & strFirstPostMsg & """>" & strSubject & "</a>")
			
			'Display the number of unread posts
			If intUnReadPostCount > 0 Then
		   		Response.Write(" [<a href=""get_last_post.asp?TID=" & lngTopicID & "&amp;ID=" & dblGetLastPostID)
		   		If intPriority = 3 Then Response.Write("&amp;FID=" & intForumID & "&amp;PR=3") 
				Response.Write("" & strQsSID2 & """ title=""" & strTxtViewUnreadPost & """>" & intUnReadPostCount & " " & strNewPostText & "</a>]") 
			End If
			
			'Display lst post details
			Response.Write("<br /><span class=""smText"">" & strTxtLastPost & ": " & strLastEntryUsername  & " " &  DateFormat(dtmLastEntryDate) & " " & strTxtAt & " " & TimeFormat(dtmLastEntryDate) & "</span>")
		
			Response.Write("</td></tr>")	


		'Else table view for all other browser
		Else


			'Write the HTML of the Topic descriptions as hyperlinks to the Topic details and message
			Response.Write(vbCrLf & " <tr class=""" & strTableRowColour & """>")
	
	
	
			'Display the status topic icons
			Response.Write(vbCrLf & "   <td align=""center"">")
			%><!-- #include file="includes/topic_status_icons_inc.asp" --><%
	     		Response.Write("</td>")
	     		
	     		
	     		
	     		
	     		Response.Write(vbCrLf & "   <td><div style=""float:left"">")
	
			 'If the user is a forum admin or a moderator then give let them delete the topic
			 If blnAdmin  OR blnModerator Then 
			 	
			 	Response.Write("<span id=""modTools" & lngTopicID & """ onclick=""showDropDown('modTools" & lngTopicID & "', 'modToolsMenu" & lngTopicID & "', 120, 0);"" class=""dropDownPointer""><img src=""" & strImagePath & "moderator_tools." & strForumImageType & """ alt=""" & strTxtModeratorTools & """ title=""" & strTxtModeratorTools & """ /></span> " & _
				"<div id=""modToolsMenu" & lngTopicID & """ class=""dropDownMenu"">" & _
				"<a href=""javascript:winOpener('pop_up_topic_admin.asp?TID=" & lngTopicID & strQsSID2 & "','admin',1,1,600,285)""><div>" & strTxtTopicAdmin & "</div></a>")
				
				'Lock or un-lock forum if admin
				If blnTopicLocked Then
					Response.Write("<a href=""lock_topic.asp?mode=UnLock&amp;TID=" & lngTopicID & "&amp;XID=" & getSessionItem("KEY") & strQsSID2 & """><div>" & strTxtUnLockTopic & "</div></a>")
				Else
					Response.Write("<a href=""lock_topic.asp?mode=Lock&amp;TID=" & lngTopicID & "&amp;XID=" & getSessionItem("KEY") & strQsSID2 & """><div>" & strTxtLockTopic & "</div></a>")
				End If
	
				'Hide or show topic
				If blnHideTopic = false Then
					Response.Write("<a href=""lock_topic.asp?mode=Hide&amp;TID=" & lngTopicID & "&amp;XID=" & getSessionItem("KEY") & strQsSID2 & """><div>" & strTxtHideTopic & "</div></a>")
				Else
					Response.Write("<a href=""lock_topic.asp?mode=Show&amp;TID=" & lngTopicID & "&amp;XID=" & getSessionItem("KEY") & strQsSID2 & """><div>" & strTxtShowTopic & "</div></a>")
				End If
				
				Response.Write("<a href=""delete_topic.asp?TID=" & lngTopicID & "&amp;PN=" & intRecordPositionPageNum & "&amp;XID=" & getSessionItem("KEY") & strQsSID2 & """ onclick=""return confirm('" & strTxtDeleteTopicAlert & "')""><div>" & strTxtDeleteTopic & "</div></a>")
				Response.Write("</div>")
			 	
			End If
			
			
			'If topic icons enabled and we have a topic icon display it	
	     		If blnTopicIcon AND strTopicIcon <> "" Then Response.Write("<img src=""" & strTopicIcon & """ alt=""" & strTxtMessageIcon & """ title=""" & strTxtMessageIcon & """ /> ")
			
			
			
			
			'Display the subject of the topic
			Response.Write("<a href=""forum_posts.asp?TID=" & lngTopicID & "&KW=" & Server.URLEncode(strSearchKeywords))
			If intPriority = 3 Then Response.Write("&FID=" & intForumID & "&PR=3")
			Response.Write("" & strQsSID2 & SeoUrlTitle(strSeoSubject, "&title=") & """ title=""" & strFirstPostMsg & """>")
			If blnBoldNewTopics AND intUnReadPostCount > 0 Then 'Unread topic subjects in bold
				Response.Write("<strong>" & strSubject & "</strong></a>")
				'Display the number of unread posts
				If intUnReadPostCount > 0 Then
		   			Response.Write(" <span class=""smText"">(<a href=""get_last_post.asp?TID=" & lngTopicID & "&amp;ID=" & dblGetLastPostID)
		   			If intPriority = 3 Then Response.Write("&amp;FID=" & intForumID & "&amp;PR=3") 
					Response.Write("" & strQsSID2 & SeoUrlTitle(strSeoSubject, "&title=") & """ title=""" & strTxtViewUnreadPost & """ class=""smLink"">" & intUnReadPostCount & " " & strNewPostText & "</a>)</span>") 
				End If
			Else
				Response.Write(strSubject & "</a>")
			End If
			
			
			'Display who started the topic and when
			Response.Write("<br /><span class=""smText"">" & strTxtBy & " <a href=""member_profile.asp?PF=" & lngTopicStartUserID & strQsSID2 & """  class=""smLink"">" & strTopicStartUsername & "</a>, " & DateFormat(dtmFirstEntryDate) & " " & strTxtAt & " " & TimeFormat(dtmFirstEntryDate) & "</span>" & _
			"</div>")
			
			
			
			'If topic rating is enabled show the rating for this topic
			If blnTopicRating AND dblTopicRating >= 1 Then
				Response.Write("<div style=""float:right;""><img src=""" & strImagePath & Mid(CStr(dblTopicRating + 0.5), 1, 1) & "_star_topic_rating." & strForumImageType & """ alt=""" & strTxtTopicRating & ": " & lngTopicVotes & " " & strTxtVotes & ", " &  strTxtAverage & " " & FormatNumber(dblTopicRating, 2) & """ title=""" & strTxtTopicRating & ": " & lngTopicVotes & " " & strTxtVotes & ", " &  strTxtAverage & " " & FormatNumber(dblTopicRating, 2) & """ /></div><br />")
			End If
	
			
			 'Calculate the number of pages for the topic and display links if there are more than 1 page
			 intNumberOfTopicPages = ((lngNumberOfReplies + 1)\intThreadsPerPage)
	
			 'If there is a remainder from calculating the num of pages add 1 to the number of pages
			 If ((lngNumberOfReplies + 1) Mod intThreadsPerPage) > 0 Then intNumberOfTopicPages = intNumberOfTopicPages + 1
	
			 'If there is more than 1 page for the topic display links to the other pages
			 If intNumberOfTopicPages > 1 Then
	
			 	Response.Write("<div style=""float:right;""><img src=""" & strImagePath & "multiple_pages." & strForumImageType & """ alt=""" & strTxtMultiplePages & """ title=""" & strTxtMultiplePages & """ />")
	
			 	'Loop round to display the links to the other pages
			 	For intTopicPagesLoopCounter = 1 To intNumberOfTopicPages
	
			 		'If there is more than 3 pages display ... last page and exit the loop
			 		If intTopicPagesLoopCounter > 3 Then
			 			
			 			'If this is position 4 then display just the 4th page
			 			If intNumberOfTopicPages = 4 Then
			 				Response.Write("<a href=""forum_posts.asp?TID=" & lngTopicID & "&KW=" & Server.URLEncode(strSearchKeywords) & "&PN=8")
							'If a priority topic need to make sure we don't change forum
				 			If intPriority = 3 Then Response.Write("&FID=" & intForumID & "&PR=3")
				 			Response.Write("" & strQsSID2 & SeoUrlTitle(strSeoSubject, "&title=") & """ class=""smPageLink"" title=""" & strTxtPage & " 4"">4</a>")
						
						'Else display the last 2 pages
						Else
						
							Response.Write("&nbsp;")
						
							Response.Write("<a href=""forum_posts.asp?TID=" & lngTopicID & "&KW=" & Server.URLEncode(strSearchKeywords) & "PN=" & intNumberOfTopicPages - 1)
							'If a priority topic need to make sure we don't change forum
				 			If intPriority = 3 Then Response.Write("&FID=" & intForumID & "&PR=3")
				 			Response.Write("" & strQsSID2 & SeoUrlTitle(strSeoSubject, "&title=") & """ class=""smPageLink"" title=""" & strTxtPage & " " & intNumberOfTopicPages - 1 & """>" & intNumberOfTopicPages - 1 & "</a>")
				 			
				 			Response.Write("<a href=""forum_posts.asp?TID=" & lngTopicID & "&KW=" & Server.URLEncode(strSearchKeywords) & "&PN=" & intNumberOfTopicPages)
							'If a priority topic need to make sure we don't change forum
				 			If intPriority = 3 Then Response.Write("&FID=" & intForumID & "&PR=3")
				 			Response.Write("" & strQsSID2 & SeoUrlTitle(strSeoSubject, "&title=") & """ class=""smPageLink"" title=""" & strTxtPage & " " & intNumberOfTopicPages & """>" & intNumberOfTopicPages & "</a>")
						
						End If
	
						'Exit the loop as we are finshed here
			 			Exit For
			 		End If
	
			 		'Display the links to the other pages
			 		Response.Write("<a href=""forum_posts.asp?TID=" & lngTopicID & "&KW=" & Server.URLEncode(strSearchKeywords) & "&PN=" & intTopicPagesLoopCounter)
			 		'If a priority topic need to make sure we don't change forum
			 		If intPriority = 3 Then Response.Write("&FID=" & intForumID & "&PR=3")
			 		Response.Write("" & strQsSID2 & SeoUrlTitle(strSeoSubject, "&title=") & """ class=""smPageLink"" title=""" & strTxtPage & " " & intTopicPagesLoopCounter & """>" & intTopicPagesLoopCounter & "</a>")
			 	Next
			 	Response.Write("</div>")
			 End If
		  %></td>
  <td align="center"><% = lngNumberOfReplies %></td>
  <td align="center"><% = lngNumberOfViews %></td>
  <td class="smText" nowrap><% = strTxtBy %> <a href="member_profile.asp?PF=<% = lngLastEntryUserID & strQsSID2 %>"  class="smLink" rel="nofollow"><% = strLastEntryUsername %></a><br/> <% = DateFormat(dtmLastEntryDate) & " " & strTxtAt & " " & TimeFormat(dtmLastEntryDate) %> <%
  		
	  		'If there are unread posts display a differnet icon and link to the last unread post
	   		If intUnReadPostCount > 0 Then
	   			 
	   			Response.Write("<a href=""get_last_post.asp?TID=" & lngTopicID & "&amp;ID=" & dblGetLastPostID & "&amp;KW=" & Server.URLEncode(strSearchKeywords))
	   			If intPriority = 3 Then Response.Write("&amp;FID=" & intForumID & "&amp;PR=3") 
				Response.Write("" & strQsSID2 & """><img src=""" & strImagePath & "view_unread_post." & strForumImageType & """ alt=""" & strTxtViewUnreadPost & " [" & intUnReadPostCount & " " & strNewPostText & "]"" title=""" & strTxtViewUnreadPost & " [" & intUnReadPostCount & " " & strNewPostText & "]"" /></a> ") 
	   	
	   		'Else there are no unread posts so display a normal topic link
	   		Else
				Response.Write("<a href=""get_last_post.asp?TID=" & lngTopicID & "&amp;ID=" & dblGetLastPostID & "&amp;KW=" & Server.URLEncode(strSearchKeywords))
				If intPriority = 3 Then Response.Write("&amp;FID=" & intForumID & "&amp;PR=3") 
				Response.Write("" & strQsSID2 & """><img src=""" & strImagePath & "view_last_post." & strForumImageType & """ alt=""" & strTxtViewLastPost & """ title=""" & strTxtViewLastPost & """ /></a> ") 
	   		End If
   
   %></td>
 </tr><%
		End If
		
		
		'Move to the next record
		intCurrentRecord = intCurrentRecord + 1
	Loop
End If


        %>
</table>
<table class="basicTable" cellspacing="0" cellpadding="4" align="center">
 <tr>
  <td><br /><!-- #include file="includes/forum_jump_inc.asp" --><%

'Release server objects
Call closeDatabase()



	%></td>
  <td align="right" valign="top" nowrap><!-- #include file="includes/page_link_inc.asp" --></td>
 </tr>
</table>
<br />
</div>
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

'Display the process time
If blnShowProcessTime Then Response.Write "<span class=""smText""><br /><br />" & strTxtThisPageWasGeneratedIn & " " & FormatNumber(Timer() - dblStartTime, 3) & " " & strTxtSeconds & "</span>"
%>
</div>
<!-- #include file="includes/footer.asp" -->