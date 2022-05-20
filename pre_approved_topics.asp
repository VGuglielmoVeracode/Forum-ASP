<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="language_files/admin_language_file_inc.asp" -->
<!--#include file="functions/functions_date_time_format.asp" -->
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
Dim dtmActiveFrom		'Holds the time to get active topics from
Dim intNumberOfTopicPages	'Holds the number of topic pages
Dim intTopicPagesLoopCounter	'Holds the number of loops
Dim blnHideTopic		'Holds if the topic is hidden
Dim strFirstPostMsg		'Holds the first posted message in the topic
Dim intForumReadRights		'Holds the read rights of the forum
Dim strForumPassword		'Holds the password for the forum
Dim strForumPaswordCode		'Holds the code for the password for the forum
Dim blnForumPasswordOK		'Set to true if the password for the forum is OK
Dim lngPollID			'Holds the topic poll id number
Dim dtmFirstEntryDate		'Holds the date of the first message
Dim intForumGroupPermission	'Holds the group permisison level for forums
Dim strTableRowColour		'Holds the row colour for the table
Dim sarryTopics			'Holds the topics to display
Dim lngTotalRecords		'Holds the number of records in the topics array
Dim lngTotalRecordsPages	'Holds the total number of pages
Dim intStartPosition		'Holds the start poition for records to be shown
Dim intEndPosition		'Holds the end poition for records to be shown
Dim intCurrentRecord		'Holds the current record position
Dim strSortDirection		'Holds the sort order
Dim strSortBy			'Holds the way the records are sorted
Dim intPageLinkLoopCounter	'Holds the loop counter for mutiple page links
Dim dtmEventDate		'Holds the date if this is a calendar event
Dim dtmEventDateEnd		'Holds the date if this is a calendar event
Dim intUnReadPostCount		'Holds the count for the number of unread posts in the forum
Dim intUnReadForumPostsLoop	'Loop to count the number of unread posts in a forum
Dim intMovedForumID
Dim dblTopicRating		'Holds the rating for a topic
Dim lngTopicVotes		'Number of votes a topic receives



'Test querystrings for any SQL Injection keywords
Call SqlInjectionTest(Request.QueryString())




'If this is the first time the page is displayed then the Forum Topic record position is set to page 1
If isNumeric(Request.QueryString("PN")) = false Then
	intRecordPositionPageNum = 1
ElseIf Request.QueryString("PN") < 1 Then
	intRecordPositionPageNum = 1

'Else the page has been displayed before so the Forum Topic record postion is set to the Record Position number
Else
	intRecordPositionPageNum = IntC(Request.QueryString("PN"))
End If



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



'Get the sort critiria
Select Case Request.QueryString("SO")
	Case "T"
		strSortBy = strDbTable & "Topic.Subject "
	Case "A"
		strSortBy = strDbTable & "Author.Username "
	Case "R"
		strSortBy = strDbTable & "Topic.No_of_replies "
	Case "V"
		strSortBy = strDbTable & "Topic.No_of_views "
	Case "UR"
		strSortBy = strDbTable & "Topic.Rating "
	Case Else
		strSortBy = strDbTable & "Topic.Last_Thread_ID "
End Select

'Sort the direction of db results
If Request.QueryString("OB") = "desc" Then
	strSortDirection = "asc"
	strSortBy = strSortBy & "DESC"
Else
	strSortDirection = "desc"
	strSortBy = strSortBy & "ASC"
End If

'If this is the first time it is run the we want dates DESC
If Request.QueryString("OB") = "" AND Request.QueryString("SO") = "" Then 
	strSortDirection = "asc"
	strSortBy = strDbTable & "Topic.Last_Thread_ID DESC"
End If




'Initilise variables
intShowTopicsFrom = 12 '12 = yesterday

'Initliase the forum group permisions
'If guest group
If intGroupID = 2 Then
	intForumGroupPermission = 1 
'If admin group
ElseIf intGroupID = 1 Then
	intForumGroupPermission = 4
'All other groups
Else
	intForumGroupPermission = 2
End If




'If new show period save to app session
If isNumeric(Request.QueryString("MT")) AND Request.QueryString("MT") <> "" Then 
	
	Call saveSessionItem("MT", Request.QueryString("MT"))
	intShowTopicsFrom = IntC(Request.QueryString("MT"))

'Get what date to show active topics
ElseIf getSessionItem("MT") <> "" Then
	
	intShowTopicsFrom = IntC(getSessionItem("MT"))

'If this is not the first time the user has visted then use year as the date to show hidden topics from
Else
	intShowTopicsFrom = 0 'All
End If




'If there is a date to show topics with then apply it to the SQL query
If intShowTopicsFrom <> 0 Then

	'Initialse the string to display when active topics are shown since
	Select Case intShowTopicsFrom
		Case 1
			strShowTopicsFrom = strTxtLastVisitOn & " " & DateFormat(dtmLastVisitDate) & " " & strTxtAt & " " & TimeFormat(dtmLastVisitDate)
			dtmActiveFrom = internationalDateTime(dtmLastVisitDate)
		case 2
			strShowTopicsFrom = strTxtLastFifteenMinutes
			dtmActiveFrom = internationalDateTime(DateAdd("n", -15, now()))
		case 3
			strShowTopicsFrom = strTxtLastThirtyMinutes
			dtmActiveFrom = internationalDateTime(DateAdd("n", -30, now()))
		Case 4
			strShowTopicsFrom = strTxtLastFortyFiveMinutes
			dtmActiveFrom = internationalDateTime(DateAdd("n", -45, now()))
		Case 5
			strShowTopicsFrom = strTxtLastHour
			dtmActiveFrom = internationalDateTime(DateAdd("h", -1, now()))
		Case 6
			strShowTopicsFrom = strTxtLastTwoHours
			dtmActiveFrom = internationalDateTime(DateAdd("h", -2, now()))
		Case 7
			strShowTopicsFrom = strTxtLastFourHours
			dtmActiveFrom = internationalDateTime(DateAdd("h", -4, now()))
		Case 8
			strShowTopicsFrom = strTxtLastSixHours
			dtmActiveFrom = internationalDateTime(DateAdd("h", -6, now()))
		Case 9
			strShowTopicsFrom = strTxtLastEightHours
			dtmActiveFrom = internationalDateTime(DateAdd("h", -8, now()))
		Case 10
			strShowTopicsFrom = strTxtLastTwelveHours
			dtmActiveFrom = internationalDateTime(DateAdd("h", -12, now()))
		Case 11
			strShowTopicsFrom = strTxtLastSixteenHours
			dtmActiveFrom = internationalDateTime(DateAdd("h", -16, now()))
		Case 12
			strShowTopicsFrom = strTxtYesterday
			dtmActiveFrom = internationalDateTime(DateAdd("d", -1, now()))
		Case 13
			strShowTopicsFrom = strTxtLastWeek
			dtmActiveFrom = internationalDateTime(DateAdd("ww", -1, now()))
		Case 14
			strShowTopicsFrom = strTxtLastMonth
			dtmActiveFrom = internationalDateTime(DateAdd("m", -1, now()))
		Case 15
			strShowTopicsFrom = strTxtLastTwoMonths
			dtmActiveFrom = internationalDateTime(DateAdd("m", -2, now()))
		Case 16
			strShowTopicsFrom = strTxtLastSixMonths
			dtmActiveFrom = internationalDateTime(DateAdd("m", -6, now()))
		Case 17
			strShowTopicsFrom = strTxtLastYear
			dtmActiveFrom = internationalDateTime(DateAdd("yyyy", -1, now()))
	End Select
End If

'If SQL server remove dash (-) from the ISO international date to make SQL Server safe
If strDatabaseType = "SQLServer" Then dtmActiveFrom = Replace(dtmActiveFrom, "-", "", 1, -1, 1)


If strDatabaseType = "Access" Then
	dtmActiveFrom = "#" & dtmActiveFrom & "#"
Else
	dtmActiveFrom = "'" & dtmActiveFrom & "'"
End If



'Initalise SQL query (quite complex but required if we only want 1 db hit to get the lot for the whole page)
strSQL = "" & _
"SELECT " & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Forum_name, " & strDbTable & "Forum.Password, " & strDbTable & "Forum.Forum_code, " & strDbTable & "Topic.Topic_ID, " & strDbTable & "Topic.Poll_ID, " & strDbTable & "Topic.Moved_ID, " & strDbTable & "Topic.Subject, " & strDbTable & "Topic.Icon, " & strDbTable & "Topic.Start_Thread_ID, " & strDbTable & "Topic.Last_Thread_ID, " & strDbTable & "Topic.No_of_replies, " & strDbTable & "Topic.No_of_views, " & strDbTable & "Topic.Locked, " & strDbTable & "Topic.Priority, " & strDbTable & "Topic.Hide, " & strDbTable & "Thread.Message_date, " & strDbTable & "Thread.Message, " & strDbTable & "Thread.Author_ID, " & strDbTable & "Author.Username, LastThread.Message_date, LastThread.Author_ID, LastAuthor.Username, " & strDbTable & "Topic.Event_date, " & strDbTable & "Topic.Event_date_end, " & strDbTable & "Topic.Rating, " & strDbTable & "Topic.Rating_Votes " & _
"FROM " & strDbTable & "Category" & strDBNoLock & ", " & strDbTable & "Forum" & strDBNoLock & ", " & strDbTable & "Topic" & strDBNoLock & ", " & strDbTable & "Thread" & strDBNoLock & ", " & strDbTable & "Thread AS LastThread" & strDBNoLock & ", " & strDbTable & "Author" & strDBNoLock & ", " & strDbTable & "Author AS LastAuthor" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Category.Cat_ID = " & strDbTable & "Forum.Cat_ID " & _
	"AND " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Topic.Forum_ID " & _
	"AND " & strDbTable & "Thread.Author_ID = " & strDbTable & "Author.Author_ID " & _
	"AND LastThread.Author_ID = LastAuthor.Author_ID " & _
	"AND " & strDbTable & "Topic.Start_Thread_ID = " & strDbTable & "Thread.Thread_ID " & _
	"AND " & strDbTable & "Topic.Last_Thread_ID = LastThread.Thread_ID "


'If not admin check they are a moderator
If blnAdmin = false Then
strSQL = strSQL & _
	"AND (" & strDbTable & "Topic.Forum_ID " & _
		"IN (" & _
			"SELECT " & strDbTable & "Permissions.Forum_ID " & _
			"FROM " & strDbTable & "Permissions " & strDBNoLock & " " & _
			"WHERE (" & strDbTable & "Permissions.Group_ID = " & intGroupID & " OR " & strDbTable & "Permissions.Author_ID = " & lngLoggedInUserID & ") " & _
				"AND " & strDbTable & "Permissions.Moderate = " & strDBTrue & _
			")" & _
	")"
End If


'Get topics with hidden posts only
strSQL = strSQL & _
	"AND (" & strDbTable & "Topic.Hide=" & strDBTrue & "  OR (" & strDbTable & "Topic.Topic_ID " & _
		"IN (" & _
			"SELECT Topic_ID " & _
			"FROM " & strDbTable & "Thread " & strDBNoLock & " " & _
			"WHERE " & strDbTable & "Thread.Hide = " & strDBTrue & _
			")" & _
		")"  & _
	")"
	

'If there is a date to show topics with then apply it to the SQL query
If intShowTopicsFrom <> 0 Then
strSQL = strSQL & _	
	"AND (LastThread.Message_date > " & dtmActiveFrom & ")"
End If




strSQL = strSQL & "ORDER BY " & strDbTable & "Category.Cat_order ASC, " & strDbTable & "Forum.Forum_Order ASC, " & strSortBy & ";"


'Set error trapping
On Error Resume Next
	
'Query the database
rsCommon.Open strSQL, adoCon

'If an error has occurred write an error to the page
If Err.Number <> 0 AND  strDatabaseType = "mySQL" Then	
	Call errorMsg("An error has occurred while executing SQL query on database.<br />Please check that the MySQL Server version is 4.1 or above.", "get_hidden_topics", "pre_approved_topics.asp")
ElseIf Err.Number <> 0 Then	
	Call errorMsg("An error has occurred while executing SQL query on database.", "get_hidden_topics", "pre_approved_topics.asp")
End If
			
'Disable error trapping
On Error goto 0

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
	

'Read in some details of the topics
If NOT rsCommon.EOF Then 
	
	'Read in the topivc recordset into an array
	sarryTopics = rsCommon.GetRows()
	
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

'Close the recordset
rsCommon.Close



'Page to link to for mutiple page (with querystrings if required)
strLinkPage = "pre_approved_topics.asp?"



'If active users is enabled update the active users application array
If blnActiveUsers Then
	'Call active users function
	saryActiveUsers = activeUsers(strTxtHiddenTopicsPosts, "", "", 0)
End If


'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & "<a href=""pre_approved_topics.asp" & strQsSID1 & """>" & strTxtHiddenTopicsPosts & "</a>"

'Status bar Active Topics Links
strStatusBarTools = strStatusBarTools & "&nbsp;&nbsp;<img src=""" & strImagePath & "active_topics." & strForumImageType & """ alt=""" & strTxtActiveTopics & """ title=""" & strTxtActiveTopics & """ style=""vertical-align: text-bottom"" /> <a href=""active_topics.asp" & strQsSID1 & """>" & strTxtActiveTopics & "</a> "
strStatusBarTools = strStatusBarTools & "&nbsp;&nbsp;<img src=""" & strImagePath & "unanswered_topics." & strForumImageType & """ alt=""" & strTxtUnAnsweredTopics & """ title=""" & strTxtUnAnsweredTopics & """ style=""vertical-align: text-bottom"" /> <a href=""active_topics.asp?UA=Y" & strQsSID2 & """>" & strTxtUnAnsweredTopics & "</a> "

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title><% = strMainForumName & " - " & strTxtHiddenTopicsPosts %></title>

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

<script  language="JavaScript">
//Function to choose how many topics are show
function ShowTopics(Show){

   	strShow = escape(Show.options[Show.selectedIndex].value);

   	if (Show != '') self.location.href = 'pre_approved_topics.asp?MT=' + strShow + '&PN=1<% = strQsSID2 %>';
	return true;
}
</script>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left">
   <h1><% = strTxtHiddenTopicsPosts %></h1>
   <span class="smText"><% = strTxtLastPostDetailNotHiddenDetails %></span>
  </td>
</table>
<br />
<table class="basicTable" cellspacing="0" cellpadding="4" align="center">
 <tr>
  <td><% = strTxtShowTopics %>
   <select name="show" id="show" onchange="ShowTopics(this)">
    <option value="0"<% If intShowTopicsFrom = 0 Then Response.Write " selected" %>><% = strTxtAnyDate %></option>
    <option value="1"<% If intShowTopicsFrom = 1 Then Response.Write " selected" %>><% = DateFormat(dtmLastVisitDate) & " " & strTxtAt & " " & TimeFormat(dtmLastVisitDate) %></option>
    <option value="2"<% If intShowTopicsFrom = 2 Then Response.Write " selected" %>><% = strTxtLastFifteenMinutes %></option>
    <option value="3"<% If intShowTopicsFrom = 3 Then Response.Write " selected" %>><% = strTxtLastThirtyMinutes %></option>
    <option value="4"<% If intShowTopicsFrom = 4 Then Response.Write " selected" %>><% = strTxtLastFortyFiveMinutes %></option>
    <option value="5"<% If intShowTopicsFrom = 5 Then Response.Write " selected" %>><% = strTxtLastHour %></option>
    <option value="6"<% If intShowTopicsFrom = 6 Then Response.Write " selected" %>><% = strTxtLastTwoHours %></option>
    <option value="7"<% If intShowTopicsFrom = 7 Then Response.Write " selected" %>><% = strTxtLastFourHours %></option>
    <option value="8"<% If intShowTopicsFrom = 8 Then Response.Write " selected" %>><% = strTxtLastSixHours %></option>
    <option value="9"<% If intShowTopicsFrom = 9 Then Response.Write " selected" %>><% = strTxtLastEightHours %></option>
    <option value="10"<% If intShowTopicsFrom = 10 Then Response.Write " selected" %>><% = strTxtLastTwelveHours %></option>
    <option value="11"<% If intShowTopicsFrom = 11 Then Response.Write " selected" %>><% = strTxtLastSixteenHours %></option>
    <option value="12"<% If intShowTopicsFrom = 12 Then Response.Write " selected" %>><% = strTxtYesterday %></option>
    <option value="13"<% If intShowTopicsFrom = 13 Then Response.Write " selected" %>><% = strTxtLastWeek %></option>
    <option value="14"<% If intShowTopicsFrom = 14 Then Response.Write " selected" %>><% = strTxtLastMonth %></option>
    <option value="15"<% If intShowTopicsFrom = 15 Then Response.Write " selected" %>><% = strTxtLastTwoMonths %></option>
    <option value="16"<% If intShowTopicsFrom = 16 Then Response.Write " selected" %>><% = strTxtLastSixMonths %></option>
    <option value="17"<% If intShowTopicsFrom = 17 Then Response.Write " selected" %>><% = strTxtLastYear %></option>
   </select>
  </td>
  <td align="right" valign="top" nowrap><!-- #include file="includes/page_link_inc.asp" --></td>
 </tr>
</table>
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
 <tr class="tableLedger">
  <td width="5%">&nbsp;</td>
  <td width="50%"><div style="float:left;"><a href="pre_approved_topics.asp?SO=T<% = strQsSID2 %>" title="<% = strTxtReverseSortOrder %>"><% = strTxtTopics %></a><% If Request.QueryString("SO") = "T" Then Response.Write(" <a href=""pre_approved_topics.asp?SO=T&OB=" & strSortDirection & strQsSID2 & """><img src=""" & strImagePath & strSortDirection & "." & strForumImageType & """ title=""" & strTxtReverseSortOrder & """ alt=""" & strTxtReverseSortOrder & """ /></a>") %> / <a href="pre_approved_topics.asp?SO=A<% = strQsSID2 %>"><% = strTxtThreadStarter %></a><% If Request.QueryString("SO") = "A" Then Response.Write(" <a href=""pre_approved_topics.asp?SO=A&OB=" & strSortDirection & strQsSID2 & """ title=""" & strTxtReverseSortOrder & """><img src=""" & strImagePath & strSortDirection & "." & strForumImageType & """ title=""" & strTxtReverseSortOrder & """ alt=""" & strTxtReverseSortOrder & """ /></a>") %></div><% 
   	'If rating is enabled
   	If blnTopicRating Then 
   		%><div style="float:right;"><a href="pre_approved_topics.asp?FID=<% = intForumID %>&amp;SO=UR<% = strQsSID2 %>&amp;OB=desc" title="<% = strTxtReverseSortOrder %>"><% = strTxtRating %></a><% If Request.QueryString("SO") = "UR" Then Response.Write(" <a href=""pre_approved_topics.asp?FID=" & intForumID & "&amp;SO=UR&amp;OB=" & strSortDirection & strQsSID2 & """><img src=""" & strImagePath & strSortDirection & "." & strForumImageType & """ alt=""" & strTxtReverseSortOrder & """ title=""" & strTxtReverseSortOrder & """ /></a>") %>&nbsp;</div><%
   		
   	End If 
   %></td>
  <td width="10%" align="center" nowrap><a href="pre_approved_topics.asp?SO=R<% = strQsSID2 %>" title="<% = strTxtReverseSortOrder %>"><% = strTxtReplies %></a><% If Request.QueryString("SO") = "R" Then Response.Write(" <a href=""pre_approved_topics.asp?SO=R&OB=" & strSortDirection & strQsSID2 & """><img src=""" & strImagePath & strSortDirection & "." & strForumImageType & """ title=""" & strTxtReverseSortOrder & """ alt=""" & strTxtReverseSortOrder & """ /></a>") %></td>
  <td width="10%" align="center" nowrap><a href="pre_approved_topics.asp?SO=V<% = strQsSID2 %>" title="<% = strTxtReverseSortOrder %>"><% = strTxtViews %></a><% If Request.QueryString("SO") = "V" Then Response.Write(" <a href=""pre_approved_topics.asp?SO=V&OB=" & strSortDirection & strQsSID2 & """><img src=""" & strImagePath & strSortDirection & "." & strForumImageType & """ title=""" & strTxtReverseSortOrder & """ alt=""" & strTxtReverseSortOrder & """ /></a>") %></td>
  <td width="20%"><a href="pre_approved_topics.asp<% = strQsSID1 %>" title="<% = strTxtReverseSortOrder %>"><% = strTxtLastPost %></a><% If Request.QueryString("SO") = "" Then Response.Write(" <a href=""pre_approved_topics.asp?OB=" & strSortDirection & strQsSID2 & """><img src=""" & strImagePath & strSortDirection & "." & strForumImageType & """ title=""" & strTxtReverseSortOrder & """ alt=""" & strTxtReverseSortOrder & """ /></a>") %></td>
 </tr><%




'If there are no active topics display an error msg
If lngTotalRecords <= 0 Then
	
	'If there are no Active Topic's to display then display the appropriate error message
	Response.Write vbCrLf & " <tr class=""tableRow""><td colspan=""6"" align=""center""><br />" & strTxtNoHiddenTopicsPostsSince & " " & strShowTopicsFrom & " " & strTxtToDisplay & "<br /><br /></td></tr>"




'Disply any active topics in the forum
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
		
		'Clean up input to prevent XXS hack
		strSubject = formatInput(strSubject)
		
		
			

		'If the forum name is different to the one from the last forum display the forum name
		If sarryTopics(1,intCurrentRecord) <> strForumName Then

			'Give the forum name the new forum name
			strForumName = sarryTopics(1,intCurrentRecord)

			'Display the new forum name
			Response.Write vbCrLf & " <tr class=""tableSubLedger""><td colspan=""7""><a href=""pre_approved_topics.asp?FID=" & intForumID & strQsSID2 & """>" & strForumName & "</a></td></tr>"
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
		End If
	
		
		
		'Calculate the row colour
		If intCurrentRecord MOD 2=0 Then strTableRowColour = "evenTableRow" Else strTableRowColour = "oddTableRow"
		
		'If this is a hidden post then change the row colour to highlight it
		If blnHideTopic Then strTableRowColour = "hiddenTableRow"




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
		Response.Write("<a href=""forum_posts.asp?TID=" & lngTopicID)
		If intPriority = 3 Then Response.Write("&FID=" & intForumID & "&PR=3")
		Response.Write("" & strQsSID2 & """ title=""" & strFirstPostMsg & """>")
		If blnBoldNewTopics AND intUnReadPostCount > 0 Then 'Unread topic subjects in bold
			Response.Write("<strong>" & strSubject & "</strong>")
		Else
			Response.Write(strSubject)
		End If
		Response.Write("</a>")

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
		 		If intTopicPagesLoopCounter > 4 Then
		 			
		 			'If this is position 4 then display just the 4th page
		 			If intNumberOfTopicPages = 4 Then
		 				Response.Write("<a href=""forum_posts.asp?TID=" & lngTopicID & "&PN=8")
						'If a priority topic need to make sure we don't change forum
			 			If intPriority = 3 Then Response.Write("&FID=" & intForumID & "&PR=3")
			 			Response.Write("" & strQsSID2 & """ class=""smPageLink"" title=""" & strTxtPage & " 4"">4</a>")
					
					'Else display the last 2 pages
					Else
					
						Response.Write("&nbsp;")
					
						Response.Write("<a href=""forum_posts.asp?TID=" & lngTopicID & "&PN=" & intNumberOfTopicPages - 1)
						'If a priority topic need to make sure we don't change forum
			 			If intPriority = 3 Then Response.Write("&FID=" & intForumID & "&PR=3")
			 			Response.Write("" & strQsSID2 & """ class=""smPageLink"" title=""" & strTxtPage & " " & intNumberOfTopicPages - 1 & """>" & intNumberOfTopicPages - 1 & "</a>")
			 			
			 			Response.Write("<a href=""forum_posts.asp?TID=" & lngTopicID & "&PN=" & intNumberOfTopicPages)
						'If a priority topic need to make sure we don't change forum
			 			If intPriority = 3 Then Response.Write("&FID=" & intForumID & "&PR=3")
			 			Response.Write("" & strQsSID2 & """ class=""smPageLink"" title=""" & strTxtPage & " " & intNumberOfTopicPages & """>" & intNumberOfTopicPages & "</a>")
					
					End If

					'Exit the loop as we are finshed here
		 			Exit For
		 		End If

		 		'Display the links to the other pages
		 		Response.Write("<a href=""forum_posts.asp?TID=" & lngTopicID & "&PN=" & intTopicPagesLoopCounter)

		 		'If a priority topic need to make sure we don't change forum
		 		If intPriority = 3 Then Response.Write("&FID=" & intForumID & "&PR=3")
		 		Response.Write("" & strQsSID2 & """ class=""smPageLink"" title=""" & strTxtPage & " " & intTopicPagesLoopCounter & """>" & intTopicPagesLoopCounter & "</a>")
		 	Next
		 	Response.Write("</div>")
		 End If
		  %></td>
  <td align="center"><% = lngNumberOfReplies %></td>
  <td align="center"><% = lngNumberOfViews %></td>
  <td class="smText" nowrap><% = strTxtBy %> <a href="member_profile.asp?PF=<% = lngLastEntryUserID & strQsSID2 %>"  class="smLink" rel="nofollow"><% = strLastEntryUsername %></a><br/> <% = DateFormat(dtmLastEntryDate) & " " & strTxtAt & " " & TimeFormat(dtmLastEntryDate) %> <%
  		
  		'If there are unread posts display a differnet icon and link to the last unread post
   		If intUnReadPostCount > 0 Then
   			 
   			Response.Write("<a href=""get_last_post.asp?TID=" & lngTopicID & "&amp;ID=" & CDbl(Now()))
   			If intPriority = 3 Then Response.Write("&amp;FID=" & intForumID & "&amp;PR=3") 
			Response.Write("" & strQsSID2 & """><img src=""" & strImagePath & "view_unread_post." & strForumImageType & """ alt=""" & strTxtViewUnreadPost & " [" & intUnReadPostCount & " " & strTxtNewPosts & "]"" title=""" & strTxtViewUnreadPost & " [" & intUnReadPostCount & " " & strTxtNewPosts & "]"" /></a> ") 
   	
   		'Else there are no unread posts so display a normal topic link
   		Else
			Response.Write("<a href=""get_last_post.asp?TID=" & lngTopicID & "&amp;ID=" & CDbl(Now()))
			If intPriority = 3 Then Response.Write("&amp;FID=" & intForumID & "&amp;PR=3") 
			Response.Write("" & strQsSID2 & """><img src=""" & strImagePath & "view_last_post." & strForumImageType & """ alt=""" & strTxtViewLastPost & """ title=""" & strTxtViewLastPost & """ /></a> ") 
   		End If
   
   %></td>
  </tr><%

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