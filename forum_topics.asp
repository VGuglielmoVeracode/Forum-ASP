<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
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
Dim intCatID			'Holds the cat ID
Dim strCatName			'Holds the cat name
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
Dim intShowTopicsFrom		'Holds when to show the topics from
Dim strShowTopicsFrom		'Holds the display text from when the topics are shown from
Dim blnForumLocked		'Set to true if the forum is locked
Dim blnTopicLocked		'set to true if the topic is locked
Dim intPriority			'Holds the priority level of the topic
Dim intNumberOfTopicPages	'Holds the number of topic pages
Dim intTopicPagesLoopCounter	'Holds the number of loops
Dim lngPollID			'Holds the Poll ID
Dim intShowTopicsWithin		'Holds the amount of time to show topics within
Dim intMovedForumID		'If the post is moved this holds the moved ID
Dim intNonPriorityTopicNum	'Holds the record count for non priority topics
Dim strFirstPostMsg		'Holds the first message in the topic
Dim dtmFirstEntryDate		'Holds the date of the first message
Dim lngLastEntryTopicID		'Holds the topic ID of the last entry
Dim strLastEntryUser		'Holds the the username of the user who made the last entry
Dim strForumDiscription		'Holds the forum description
Dim strForumPassword		'Holds the forum password if there is one
Dim lngNumberOfTopics		'Holds the number of topics in a forum
Dim lngNumberOfPosts		'Holds the number of Posts in the forum
Dim blnHideForum		'Set to true if this is a hidden forum
Dim intForumColourNumber	'Holds the number to calculate the table row colour
Dim lngLastEntryMeassgeID	'Holds the message ID of the last entry
Dim intMasterForumID		'Holds the main forum ID
Dim strMasterForumName		'Holds the main forum name
Dim blnHideTopic		'Set to true if the topic is hidden
Dim strTableRowColour		'Holds the row colour for the table
Dim sarryTopics			'Holds the topics to display
Dim sarrySubForums		'Holds the sub forums to display
Dim intSubForumID		'Holds the sub forum ID
Dim intCurrentRecord		'Holds the current record position
Dim lngTotalRecords		'Holds the number of records in the topics array
Dim lngTotalRecordsPages	'Holds the total number of pages
Dim intStartPosition		'Holds the start poition for records to be shown
Dim intEndPosition		'Holds the end poition for records to be shown
Dim intSubCurrentRecord		'Holds the current records for the sub forums
Dim strSortDirection		'Holds the sort order
Dim strSortBy			'Holds the way the records are sorted
Dim strShowTopicsDate		'Holds the show topics date
Dim dtmEventDate		'Holds the date if this is a calendar event
Dim dtmEventDateEnd		'Holds the date if this is a calendar event
Dim intPageLinkLoopCounter
Dim intUnReadPostCount		'Holds the count for the number of unread posts in the forum
Dim intUnReadForumPostsLoop	'Loop to count the number of unread posts in a forum
Dim intMaxResults		'Max results retrurned
Dim strSubForums		'Holds the subforums names
Dim strGroupTag			'Holds the group tag
Dim strForumDescription		'Used for meta tags
Dim strSQLFromWhere		'Used for the From where clause
Dim dblTopicRating		'Holds the rating for a topic
Dim lngTopicVotes		'Number of votes a topic receives
Dim strDynamicKeywords
Dim strNewPostText
Dim strPageQueryString		'Holds the querystring for the page	
Dim strCanonicalURL
Dim strForumImageIcon		'Hold an image icon for the forum
Dim strForumURL 		'Holds the forum URL if a link
Dim dblGetLastPostID


'Initlaise variables
intNonPriorityTopicNum = 0
blnHideTopic = false
intSubCurrentRecord = 0
intCurrentRecord = 0
strShowTopicsFrom = " '" & strTxtAnyDate & "'"
dblGetLastPostID = CDbl(Now())


'Read in the qerystring
strPageQueryString = Request.QueryString()


'Remove the page title from the querystring beofre doing the sql injection test
If Request.QueryString("title") <> "" Then strPageQueryString = Replace(strPageQueryString, Request.QueryString("title"), "")

'Test querystrings for any SQL Injection keywords
Call SqlInjectionTest(strPageQueryString)



'Max results database returned reseults
'Only used if database doesn't support server side paging or if paging disabled
If strDatabaseType = "Access" Then
	intMaxResults = 500
Else
	intMaxResults = 5000
End If



'Read in the Forum ID to display the Topics for
If isNumeric(Request.QueryString("FID")) Then intForumID = IntC(Request.QueryString("FID")) Else intForumID = 0



'If there is no Forum ID to display the Topics for then redirect the user to the main forum page
If intForumID = 0 Then

	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("default.asp" & strQsSID1)
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

'Calculate the start position of the records to get from db
intStartPosition = ((intRecordPositionPageNum - 1) * intTopicPerPage)




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




'Read in the forum details inc. cat name, forum details, and permissions (also reads in the main forum name if in a sub forum, saves on db call later)

'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "" & _
"SELECT" & strDBTop1 & " " & strDbTable & "Category.Cat_ID, " & strDbTable & "Category.Cat_name, " & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Sub_ID, " & strDbTable & "Forum.Forum_name, " & strDbTable & "Forum.Forum_description, " & strDbTable & "Forum2.Forum_name AS Main_forum, " & strDbTable & "Forum.Password, " & strDbTable & "Forum.Forum_code, " & strDbTable & "Forum.Locked, " & strDbTable & "Forum.Show_topics, " & strDbTable & "Forum.Forum_URL, " & strDbTable & "Permissions.* " & _
"FROM " & strDbTable & "Category" & strDBNoLock & ", " & strDbTable & "Forum" & strDBNoLock & ", " & strDbTable & "Forum AS " & strDbTable & "Forum2" & strDBNoLock & ", " & strDbTable & "Permissions" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Category.Cat_ID = " & strDbTable & "Forum.Cat_ID " & _
	"AND " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Permissions.Forum_ID " & _
	"AND " & strDbTable & "Forum.Forum_ID = " & intForumID & " " & _
	"AND (" & strDbTable & "Forum.Sub_ID = " & strDbTable & "Forum2.Forum_ID OR (" & strDbTable & "Forum.Sub_ID = 0 AND " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Forum2.Forum_ID)) " & _
	"AND (" & strDbTable & "Permissions.Author_ID = " & lngLoggedInUserID & " OR " & strDbTable & "Permissions.Group_ID = " & intGroupID & ") " & _
"ORDER BY " & strDbTable & "Forum.Sub_ID, " & strDbTable & "Forum.Forum_Order, " & strDbTable & "Permissions.Author_ID DESC" & strDBLimit1 & ";"


'Set error trapping
On Error Resume Next

'Query the database
rsCommon.Open strSQL, adoCon

'If an error has occurred write an error to the page
If Err.Number <> 0 Then	Call errorMsg("An error has occurred while executing SQL query on database.", "get_forum_data", "forum_topics.asp")

'Disable error trapping
On Error goto 0

'If there is a record returned by the recordset then check to see if you need a password to enter it
If NOT rsCommon.EOF Then

	'Read in forum details from the database
	intCatID = CInt(rsCommon("Cat_ID"))
	strCatName = rsCommon("Cat_name")
	strForumName = rsCommon("Forum_name")
	strMasterForumName = rsCommon("Main_forum")
	intMasterForumID = CLng(rsCommon("Sub_ID"))
	blnForumLocked = CBool(rsCommon("Locked"))
	intShowTopicsWithin = CInt(rsCommon("Show_topics"))
	strForumDescription = removeHTML(rsCommon("Forum_description"), 255, True)
	strForumURL = rsCommon("Forum_URL")
	
	'If forum URL just has http:// then blank it 
	If strForumURL = "http://" OR isNull(strForumURL) Then strForumURL = ""

	'Read in the forum permissions
	blnRead = CBool(rsCommon("View_Forum"))
	blnPost = CBool(rsCommon("Post"))
	blnReply = CBool(rsCommon("Reply_posts"))
	blnEdit = CBool(rsCommon("Edit_posts"))
	blnDelete = CBool(rsCommon("Delete_posts"))
	blnPriority = CBool(rsCommon("Priority_posts"))
	blnPollCreate = CBool(rsCommon("Poll_create"))
	blnVote = CBool(rsCommon("Vote"))
	blnModerator = CBool(rsCommon("Moderate"))
	blnCheckFirst = CBool(rsCommon("Display_post"))
	
	'If the title is not included in the link redirect page back to itself with the title included
	If Request.QueryString("title") <> "" AND Request.QueryString("title") <> SeoUrlTitle(strForumName, "")  Then
		
		'Clean up
		rsCommon.Close
		Call closeDatabase()
	
		'Redirect back to page with correct page name
		Response.Redirect("forum_topics.asp?FID=" & intForumID & SeoUrlTitle(strForumName, "&title="))
		Response.End
		
	End If	

	'If the user has no read writes then kick them
	If blnRead = False Then

		'Reset Server Objects
		rsCommon.Close
		Call closeDatabase()

		'Redirect to a page asking for the user to enter the forum password
		Response.Redirect("insufficient_permission.asp" & strQsSID1)
	End If
	
	
	'If a forum URL then is is a forum link so redirect
	If strForumURL <> "" Then
		
		'Reset Server Objects
		rsCommon.Close
		Call closeDatabase()
		
		'Redirect to forum link
		Response.Redirect(strForumURL)
		
	End If


	'If the forum requires a password and a logged in forum code is not found on the users machine then send them to a login page
	If rsCommon("Password") <> "" AND (getCookie("fID", "Forum" & intForumID) <> rsCommon("Forum_code") AND getSessionItem("FP" & intForumID) <> "1") Then

		'Reset Server Objects
		rsCommon.Close
		Call closeDatabase()


		'Redirect to a page asking for the user to enter the forum password
		Response.Redirect("forum_password_form.asp?FID=" & intForumID & strQsSID3)
	End If
	
End If

'Clean up
rsCommon.Close



'Get what date to show topics till from querystring
'If mobile browser then always set to any date as mobile browser does not have the date select drop down
If blnMobileBrowser Then
	
	intShowTopicsFrom = 0 
	
ElseIf isNumeric(Request.QueryString("TS")) AND Request.QueryString("TS") <> "" Then

	Call saveSessionItem("TS", Request.QueryString("TS"))
	intShowTopicsFrom = IntC(Request.QueryString("TS"))

'Get what date to show topics
ElseIf getSessionItem("TS") <> "" Then

	intShowTopicsFrom = IntC(getSessionItem("TS"))

'Else if there is no cookie use the default set by the forum
Else
	intShowTopicsFrom = intShowTopicsWithin
End If


'Get the sort critiria
Select Case UCase(Request.QueryString("SO"))
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






'********************************
'****   Get Topics SQL Query ****
'********************************


'Start with WHERE Cluases as these are used in both the count query and in the main query
'If there is a date to show topics with then apply it to the SQL query
If intShowTopicsFrom <> 0 Then

	strSQLFromWhere = "AND ((LastThread.Message_date > "

	'If Access use # around dates, other DB's use ' around dates
	If strDatabaseType = "Access" Then
		strSQLFromWhere = strSQLFromWhere & "#"
	Else
		strSQLFromWhere = strSQLFromWhere & "'"
	End If

	'Initialse the string to display when active topics are shown since
	Select Case intShowTopicsFrom
		Case 1
			strShowTopicsFrom = strTxtLastVisitOn & " " & DateFormat(dtmLastVisitDate) & " " & strTxtAt & " " & TimeFormat(dtmLastVisitDate)
			strShowTopicsDate = internationalDateTime(dtmLastVisitDate)
		Case 2
			strShowTopicsFrom = strTxtYesterday
			strShowTopicsDate = internationalDateTime(DateAdd("d", -1, now()))
		Case 3
			strShowTopicsFrom = strTxtLastTwoDays
			strShowTopicsDate = internationalDateTime(DateAdd("d", -2, now()))
		Case 4
			strShowTopicsFrom = strTxtLastWeek
			strShowTopicsDate = internationalDateTime(DateAdd("ww", -1, now()))
		Case 5
			strShowTopicsFrom = strTxtLastMonth
			strShowTopicsDate = internationalDateTime(DateAdd("m", -1, now()))
		Case 6
			strShowTopicsFrom = strTxtLastTwoMonths
			strShowTopicsDate = internationalDateTime(DateAdd("m", -2, now()))
		Case 7
			strShowTopicsFrom = strTxtLastSixMonths
			strShowTopicsDate = internationalDateTime(DateAdd("m", -6, now()))
		Case 8
			strShowTopicsFrom = strTxtLastYear
			strShowTopicsDate = internationalDateTime(DateAdd("yyyy", -1, now()))
	End Select


	'If SQL server remove dash (-) from the ISO international date to make SQL Server safe
	If strDatabaseType = "SQLServer" Then strShowTopicsDate = Replace(strShowTopicsDate, "-", "", 1, -1, 1)

	'Place into SQL query
	strSQLFromWhere = strSQLFromWhere & strShowTopicsDate

	'If Access use # around dates, other DB's use ' around dates
	If strDatabaseType = "Access" Then
		strSQLFromWhere = strSQLFromWhere & "#"
	Else
		strSQLFromWhere = strSQLFromWhere & "'"
	End If

	strSQLFromWhere = strSQLFromWhere & ") OR (" & strDbTable & "Topic.Priority > 0)) "
	
	
	
End If

'Select which topics to get
strSQLFromWhere = strSQLFromWhere & "AND (" & strDbTable & "Topic.Priority = 3 " & _
		"OR " & strDbTable & "Topic.Moved_ID = " & intForumID & " " & _
		"OR " & strDbTable & "Topic.Forum_ID = " & intForumID & ") " & _
	") "

'If this isn't a moderator only display hidden posts if the user posted them
If blnModerator = false AND blnAdmin = false Then
	strSQLFromWhere = strSQLFromWhere & "AND (" & strDbTable & "Topic.Hide = " & strDBFalse & " "
	'Don't display hidden posts if guest
	If intGroupID <> 2 Then strSQLFromWhere = strSQLFromWhere & "OR " & strDbTable & "Thread.Author_ID = " & lngLoggedInUserID
	strSQLFromWhere = strSQLFromWhere & ") "
End If





'If using advanced paging then we need to count the total number of records
If (strDatabaseType = "SQLServer" AND blnSqlSvrAdvPaging) OR strDatabaseType = "mySQL" Then
	
	strSQL = "" & _
	"SELECT Count(" & strDbTable & "Topic.Topic_ID) AS TopicCount " & _
	"FROM " & strDbTable & "Topic" & strDBNoLock & ", " & strDbTable & "Thread" & strDBNoLock & " , " & strDbTable & "Thread AS LastThread " & _
	"WHERE  (" & strDbTable & "Topic.Last_Thread_ID = " & strDbTable & "Thread.Thread_ID " & _
		"AND " & strDbTable & "Topic.Last_Thread_ID = LastThread.Thread_ID "
	
	strSQL = strSQL & strSQLFromWhere & ";"

	'Set error trapping
	On Error Resume Next
		
	'Query the database
	rsCommon.Open strSQL, adoCon
	
	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then	Call errorMsg("An error has occurred while executing SQL query on database.", "topic_count", "forum_topics.asp")
				
	'Disable error trapping
	On Error goto 0
	
	'Read in member count from database
	lngTotalRecords = CLng(rsCommon("TopicCount"))
	
	'Close recordset
	rsCommon.close
End If






'Read in all the topics for this forum and place them in an array
strSQL = "" & _
"SELECT "

'If SQL server advanced paging
If strDatabaseType = "SQLServer" AND blnSqlSvrAdvPaging Then
	strSQL = strSQL & " * " & _
	"FROM (SELECT "
	'Using TOP really speeds thing up on the first pages, but once you get up to really large number of pages it can cause order problems
	If intThreadsPerPage * intRecordPositionPageNum < 500 Then strSQL = strSQL & " TOP " & intTopicPerPage * intRecordPositionPageNum  & " "
	
'Access is naff, no database paging so just limit the results (same if paging is disabled)
ElseIf strDatabaseType = "Access" OR strDatabaseType = "SQLServer" Then
	strSQL = strSQL & " TOP " & intMaxResults & " "
End If

'Rest of query
strSQL = strSQL & _
" " & strDbTable & "Topic.Topic_ID, " & strDbTable & "Topic.Poll_ID, " & strDbTable & "Topic.Moved_ID, " & strDbTable & "Topic.Subject, " & strDbTable & "Topic.Icon, " & strDbTable & "Topic.Start_Thread_ID, " & strDbTable & "Topic.Last_Thread_ID, " & strDbTable & "Topic.No_of_replies, " & strDbTable & "Topic.No_of_views, " & strDbTable & "Topic.Locked, " & strDbTable & "Topic.Priority, " & strDbTable & "Topic.Hide, " & strDbTable & "Thread.Message_date, " & strDbTable & "Thread.Message, " & strDbTable & "Thread.Author_ID, " & strDbTable & "Author.Username, LastThread.Message_date AS LastMessageDate, LastThread.Author_ID AS LastTreadDate, LastAuthor.Username AS LastUsername, " & strDbTable & "Topic.Event_date, " & strDbTable & "Topic.Event_date_end, " & strDbTable & "Topic.Rating, " & strDbTable & "Topic.Rating_Votes "

'If SQL Server advanced paging
If strDatabaseType = "SQLServer"  AND blnSqlSvrAdvPaging Then
	strSQL = strSQL & ", ROW_NUMBER() OVER (ORDER BY " & strDbTable & "Topic.Priority DESC, " & strSortBy & ") AS RowNum "

End If


strSQL = strSQL & "" & _
"FROM " & strDbTable & "Topic" & strDBNoLock & ", " & strDbTable & "Thread" & strDBNoLock & ", " & strDbTable & "Thread AS LastThread" & strDBNoLock & ", " & strDbTable & "Author" & strDBNoLock & ", " & strDbTable & "Author AS LastAuthor" & strDBNoLock & " " & _
"WHERE ("

'Do the table joins
strSQL = strSQL & " " & strDbTable & "Thread.Author_ID = " & strDbTable & "Author.Author_ID " & _
	"AND LastThread.Author_ID = LastAuthor.Author_ID " & _
	"AND " & strDbTable & "Topic.Start_Thread_ID = " & strDbTable & "Thread.Thread_ID " & _
	"AND " & strDbTable & "Topic.Last_Thread_ID = LastThread.Thread_ID "

'Get the WHERE Clouses setup earlier
strSQL = strSQL & strSQLFromWhere

'If SQL Server advanced paging
If strDatabaseType = "SQLServer" AND blnSqlSvrAdvPaging Then
	strSQL = strSQL & ") AS PagingQuery WHERE RowNum BETWEEN " & intStartPosition + 1 & " AND " & intStartPosition + intTopicPerPage & " "

'Else Order by clause here
Else
	strSQL = strSQL & "ORDER BY " & strDbTable & "Topic.Priority DESC, " & strSortBy & " "
End If

'mySQL limit operator
If strDatabaseType = "mySQL" Then
	strSQL = strSQL & " LIMIT " & intStartPosition & ", " & intTopicPerPage
End If

strSQL = strSQL & ";"



'Set error trapping
On Error Resume Next

'Query the database
rsCommon.Open strSQL, adoCon

'If an error has occurred write an error to the page
If Err.Number <> 0 Then	Call errorMsg("An error has occurred while executing SQL query on database.", "get_topics_data", "forum_topics.asp")

'Disable error trapping
On Error goto 0




'SQL Query Array Look Up table
'0 = tblTopic.Topic_ID
'1 = tblTopic.Poll_ID
'2 = tblTopic.Moved_ID
'3 = tblTopic.Subject
'4 = tblTopic.Icon
'5 = tblTopic.Start_Thread_ID
'6 = tblTopic.Last_Thread_ID
'7 = tblTopic.No_of_replies
'8 = tblTopic.No_of_views
'9 = tblTopic.Locked
'10 = tblTopic.Priority
'11 = tblTopic.Hide
'12 = tblThread.Message_date
'13 = tblThread.Message,
'14 = tblThread.Author_ID,
'15 = tblAuthor.Username,
'16 = LastThread.Message_date,
'17 = LastThread.Author_ID,
'18 = LastAuthor.Username
'19 = tblTopic.Event_date
'20 = tblTopic.Event_date_end
'21 = tblTopic.Rating
'22 = tblTopic.Rating_Votes



'Read in some details of the topics
If NOT rsCommon.EOF Then

	'Read in the Topic recordset into an array
	sarryTopics = rsCommon.GetRows()
	
	'If advanced paging then workout the end and start position differently
	If (strDatabaseType = "SQLServer" AND blnSqlSvrAdvPaging) OR strDatabaseType = "mySQL" Then
		
		'End Position
		intEndPosition = Ubound(sarryTopics,2) + 1
	
		'Get the start position
		intCurrentRecord = 0
	
	'Else standard slower paging	
	Else
		'Count the number of records
		lngTotalRecords = Ubound(sarryTopics,2) + 1
	
		'Start position
		intStartPosition = ((intRecordPositionPageNum - 1) * intTopicPerPage)
	
		'End Position
		intEndPosition = intStartPosition + intTopicPerPage
	
		'Get the start position
		intCurrentRecord = intStartPosition
	End If
	
	'Count the number of pages for the topics using FIX so that we get the whole number and  not any fractions
	lngTotalRecordsPages = FIX(lngTotalRecords / intTopicPerPage)
	
	'If there is a remainder or the result is 0 then add 1 to the total num of pages
	If lngTotalRecords Mod intTopicPerPage > 0 OR lngTotalRecordsPages = 0 Then lngTotalRecordsPages = lngTotalRecordsPages + 1
	
End If

'Close the recordset
rsCommon.Close




'********************************
'**** 	 Sub forums	     ****
'********************************

'Read the various forums from the database
'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "" & _
"SELECT " & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Forum_name, " & strDbTable & "Forum.Forum_description, " & strDbTable & "Forum.No_of_topics, " & strDbTable & "Forum.No_of_posts, " & strDbTable & "Author.Username, " & strDbTable & "Forum.Last_post_author_ID, " & strDbTable & "Forum.Last_post_date, " & strDbTable & "Forum.Password, " & strDbTable & "Forum.Locked, " & strDbTable & "Forum.Hide, " & strDbTable & "Permissions.View_Forum, " & strDbTable & "Forum.Last_topic_ID, " & strDbTable & "Topic.Subject, " & strDbTable & "Forum.Forum_icon " & _
"FROM (((" & strDbTable & "Category INNER JOIN " & strDbTable & "Forum ON " & strDbTable & "Category.Cat_ID = " & strDbTable & "Forum.Cat_ID) LEFT JOIN " & strDbTable & "Topic ON " & strDbTable & "Forum.Last_topic_ID = " & strDbTable & "Topic.Topic_ID) INNER JOIN " & strDbTable & "Author ON " & strDbTable & "Forum.Last_post_author_ID = " & strDbTable & "Author.Author_ID) INNER JOIN " & strDbTable & "Permissions ON " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Permissions.Forum_ID " & _
"WHERE " & strDbTable & "Forum.Sub_ID = " & intForumID & " " & _
	"AND (" & strDbTable & "Permissions.Author_ID = " & lngLoggedInUserID & " OR " & strDbTable & "Permissions.Group_ID = " & intGroupID & ") " & _
"ORDER BY " & strDbTable & "Forum.Forum_Order, " & strDbTable & "Permissions.Forum_ID;"


'Set error trapping
On Error Resume Next

'Query the database
rsCommon.Open strSQL, adoCon

'If an error has occurred write an error to the page
If Err.Number <> 0 Then	Call errorMsg("An error has occurred while executing SQL query on database.", "get_sub_forum_data", "forum_topics.asp")

'Disable error trapping
On Error goto 0

'If there are sub forums to dsiplay, then display them
If NOT rsCommon.EOF Then

	'Read in the sub forum details into an array
	sarrySubForums = rsCommon.GetRows()
End If

'Close the recordset
rsCommon.Close





'Use the application session to pass around what forum this user is within
If IntC(getSessionItem("FID")) <> intForumID Then Call saveSessionItem("FID", intForumID)



'Page to link to for mutiple page (with querystrings if required)
strLinkPage = "forum_topics.asp?FID=" & intForumID & "&amp;"
strLinkPageTitle = SeoUrlTitle(strForumName, "&amp;title=")




'If active users is enabled update the active users application array
If blnActiveUsers Then
	'Call active users function
	saryActiveUsers = activeUsers(strTxtViewingIndex, strForumName, "forum_topics.asp?FID=" & intForumID & SeoUrlTitle(strForumName, "&amp;title="), 0)
End If



'If URL Rewriting is enabled create the canonical to the page for improved SEO
If NOT Request.ServerVariables("HTTP_X_ORIGINAL_URL") = "" OR (NOT Request.ServerVariables("HTTP_X_REWRITE_URL") = "" AND InStr(Request.ServerVariables("HTTP_X_REWRITE_URL"), ".html") > 1) Then
	
	If intRecordPositionPageNum = 1 Then
		strCanonicalURL = strForumPath & SeoUrlTitle(strForumName, "") & "_forum" & intForumID & ".html"
	Else
		strCanonicalURL = strForumPath & SeoUrlTitle(strForumName, "") & "_forum" & intForumID & "_page" & intRecordPositionPageNum & ".html"
	End If

'Else canonical without URL rewriting
Else
	If intRecordPositionPageNum = 1 Then
		strCanonicalURL = strForumPath & "forum_topics.asp?FID=" & intForumID & SeoUrlTitle(strForumName, "&title=")
	Else
		strCanonicalURL = strForumPath & "forum_topics.asp?FID=" & intForumID & "&PN=" & intRecordPositionPageNum & SeoUrlTitle(strForumName, "&title=")
	End If	
End If



'Set bread crumb trail
'Display the category name
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & "<a href=""default.asp?C=" & intCatID & strQsSID2 & SeoUrlTitle(strCatName, "&title=") & """>" & strCatName & "</a>" & strNavSpacer

'Display if there is a main forum to the sub forums name
If intMasterForumID <> 0 Then strBreadCrumbTrail = strBreadCrumbTrail & "<a href=""forum_topics.asp?FID=" & intMasterForumID & strQsSID2 & SeoUrlTitle(strMasterForumName, "&title=") & """>" & strMasterForumName & "</a>" & strNavSpacer

'Display forum name
If strForumName = "" Then strBreadCrumbTrail = strBreadCrumbTrail &  strTxtNoForums Else strBreadCrumbTrail = strBreadCrumbTrail & "<a href=""forum_topics.asp?FID=" & intForumID & strQsSID2 & SeoUrlTitle(strForumName, "&title=") & """>" & strForumName & "</a>"






'Set the status bar tools

'Modertor Tools
If blnAdmin OR blnModerator Then
	strStatusBarTools = strStatusBarTools & "&nbsp;&nbsp;<span id=""modTools"" onclick=""showDropDown('modTools', 'modToolsMenu', 120, 0);"" class=""dropDownPointer""><img src=""" & strImagePath & "moderator_tools." & strForumImageType & """ alt=""" & strTxtModeratorTools & """ title=""" & strTxtModeratorTools & """ style=""vertical-align: text-bottom"" /> " & strTxtModeratorTools & "</span>" & _
	"<div id=""modToolsMenu"" class=""dropDownMenu"">" & _
	"<a href=""pre_approved_topics.asp" & strQsSID1 & """><div>" & strTxtHiddenTopics & "</div></a>" & _
	"<a href=""resync_post_count.asp?FID=" & intForumID & strQsSID2 & """><div>" & strTxtResyncTopicPostCount & "</div></a>"
	
	'Lock or un-lock forum if admin
	If blnAdmin AND blnForumLocked Then
		strStatusBarTools = strStatusBarTools & "<a href=""lock_forum.asp?mode=UnLock&amp;FID=" & intForumID & "&amp;XID=" & getSessionItem("KEY") & strQsSID2 & """><div>" & strTxtUnForumLocked & "</div></a>"
	Else
		strStatusBarTools = strStatusBarTools & "<a href=""lock_forum.asp?mode=Lock&amp;FID=" & intForumID & "&amp;XID=" & getSessionItem("KEY") & strQsSID2 & """><div>" & strTxtLockForum & "</div></a>"
	End If
	
	strStatusBarTools = strStatusBarTools & "</div>"
End If


'Topic Options drop down
strStatusBarTools = strStatusBarTools & "&nbsp;&nbsp;<span id=""forumOptions"" onclick="""
'If we need a subscription link then include a call to the ajax function
If blnEmail AND blnLoggedInUserEmail AND intGroupID <> 2 AND blnActiveMember Then strStatusBarTools = strStatusBarTools & "getAjaxData('ajax_email_notify.asp?FID=" & intForumID & "&amp;PN=" & intRecordPositionPageNum & strQsSID2 & "', 'ajaxEmailSub');"
	
strStatusBarTools = strStatusBarTools & "showDropDown('forumOptions', 'optionsMenu', 132, 0);"" class=""dropDownPointer""><img src=""" & strImagePath & "forum_options." & strForumImageType & """ alt=""" & strTxtForumOptions & """ title=""" & strTxtForumOptions & """ style=""vertical-align: text-bottom"" /> <a href=""#"">" & strTxtForumOptions & "</a></span>" & _
"<div id=""optionsMenu"" class=""dropDownMenu"">" & _
"<a href=""new_topic_form.asp?FID=" & intForumID & strQsSID2 & """><div>" & strTxtCreateNewTopic & "</div></a>"

'If the user can create a poll disply a create poll link
If blnPollCreate Then strStatusBarTools = strStatusBarTools & "<a href=""new_poll_form.asp?FID=" & intForumID & strQsSID2 & """><div>" & strTxtCreateNewPoll & "</div></a>"

'Display option to subscribe or un-subscribe to forum
If blnEmail AND blnLoggedInUserEmail AND intGroupID <> 2 AND blnActiveMember Then strStatusBarTools = strStatusBarTools & "<span id=""ajaxEmailSub""></span>"

strStatusBarTools = strStatusBarTools & "</div>"


'Active Topics Links
strStatusBarTools = strStatusBarTools & "&nbsp;&nbsp;<img src=""" & strImagePath & "active_topics." & strForumImageType & """ alt=""" & strTxtActiveTopics & """ title=""" & strTxtActiveTopics & """ style=""vertical-align: text-bottom"" /> <a href=""active_topics.asp" & strQsSID1 & """>" & strTxtActiveTopics & "</a>"

'If RSS XML enabled then display an RSS button to link to XML file
If blnRSS Then strStatusBarTools = strStatusBarTools & " <a href=""RSS_topic_feed.asp?FID=" & intForumID & SeoUrlTitle(strForumName, "&title=") & """ target=""_blank""><img src=""" & strImagePath & "rss." & strForumImageType & """ alt=""" & strTxtRSS & ": " & strForumName & """ title=""" & strTxtRSS & " - " & strForumName & """ /></a>"



%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title><% = strForumName & " - " & strMainForumName  %><% If lngTotalRecordsPages > 1 Then Response.Write(" - " & strTxtPage & " " & intRecordPositionPageNum) %></title>
<meta name="generator" content="Web Wiz Forums <% = strVersion %>" /><% 

'Dynamic meta tags
If blnDynamicMetaTags Then 
%>
<meta name="description" content="<% = strForumDescription %>" />
<meta name="keywords" content="<% = dynamicKeywords(strForumDescription) & ", " & strBoardMetaKeywords %>" /><%

Else
	
%>
<meta name="description" content="<% = strBoardMetaDescription %>" />
<meta name="keywords" content="<% = strBoardMetaKeywords %>" /><%

End If


'Display Canonical URL Meta tag
If NOT strCanonicalURL = "" Then Response.Write(vbCrLf & "<link rel=""canonical"" href=""" & strCanonicalURL & """ />")
	
	
	

'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write(vbCrLf & vbCrLf & "<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Forums(TM) ver. " & strVersion & "" & _
vbCrLf & "Info: http://www.webwizforums.com" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******


'If RSS feed is enabled then have an alt link to it for browers that support RSS Feeds
If blnRSS Then Response.Write(vbCrLf & "<link rel=""alternate"" type=""application/rss+xml"" title=""RSS 2.0"" href=""RSS_topic_feed.asp?FID=" & intForumID & SeoUrlTitle(strForumName, "&title=") & """ />")
%>
<script language="JavaScript" type="text/javascript">
function ShowTopics(Show){
   	strShow = escape(Show.options[Show.selectedIndex].value);
   	if (Show != '') self.location.href = 'forum_topics.asp?FID=<% = intForumID %>&TS=' + strShow + '<% = strQsSID2 %>';
	return true;
}
</script>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><%

'If the forum is locked show a locked pad lock icon
If blnForumLocked = True Then Response.Write ("<img src=""" & strImagePath & "forum_lock." & strForumImageType & """ alt=""" & strTxtForumLocked & """ title=""" & strTxtForumLocked & """ style=""vertical-align: text-bottom"" />")

'Display the forum name
Response.Write(strForumName)

%></h1></td>
</tr>
</table><%


'******************************************
'***	      Display Sub Forums	***
'******************************************

'If there are sub forums to dsiplay, then display them
If isArray(sarrySubForums) Then


	'SQL Query Array Look Up table
	'0 = Forum_ID
	'1 = Forum_name
	'2 = Forum_description
	'3 = No_of_topics
	'4 = No_of_posts
	'5 = Username
	'6 = Last_post_author_ID
	'7 = Last_post_date
	'8 = Password
	'9 = Locked
	'10 = Hide
	'11 = Read
	'12 = Last_topic_ID
	'13 = Topic.Subject
	'14 = Forum_icon


	'First we need to loop through the array to see if the user can view any of the sub forums if not don't display them
	'If we find a sub forum that the user can view we exit this loop to display the sub forum
	Do While intSubCurrentRecord <= Ubound(sarrySubForums,2)

		'Read in details
		blnHideForum = CBool(sarrySubForums(10,intSubCurrentRecord))
		blnRead = CBool(sarrySubForums(11,intSubCurrentRecord))

		'If this forum is to be shown then leave the loop and display the sub forums
		If blnHideForum = False OR blnRead Then Exit Do

		'Move to next record
		intSubCurrentRecord = intSubCurrentRecord + 1
	Loop


	'If there are still records left in the array then these are the forums that the user can view so display them
	If intSubCurrentRecord <= Ubound(sarrySubForums,2) Then


%>
<br />
<table class="tableBorder" align="center" cellspacing="1" cellpadding="10"><%

		'Display single column heading for mobile users
		If blnMobileBrowser Then
%>
 <tr class="tableLedger">
  <td><% = strTxtSub & " " & strTxtForums %></td>
 </tr><%
		
		'Display all column headings
		Else
%>
 <tr class="tableLedger">
  <td width="5%">&nbsp;</td>
  <td width="50%"><% = strTxtSub & " " & strTxtForums %></td>
  <td width="10%" align="center"><% = strTxtTopics %></td>
  <td width="10%" align="center"><% = strTxtPosts %></td>
  <td width="30%"><% = strTxtLastPost %></td>
 </tr><%
		End If

		'Loop through the array and display the forums
		Do While intSubCurrentRecord <= Ubound(sarrySubForums,2)

			'Initialise variables
			lngLastEntryTopicID = 0

			'Read in the details for this forum
			intSubForumID = CInt(sarrySubForums(0,intSubCurrentRecord))
			strForumName = sarrySubForums(1,intSubCurrentRecord)
			strForumDiscription = sarrySubForums(2,intSubCurrentRecord)
			lngNumberOfTopics = CLng(sarrySubForums(3,intSubCurrentRecord))
			lngNumberOfPosts = CLng(sarrySubForums(4,intSubCurrentRecord))
			strLastEntryUser = sarrySubForums(5,intSubCurrentRecord)
			If isNumeric(sarrySubForums(6,intSubCurrentRecord)) Then lngLastEntryUserID = CLng(sarrySubForums(6,intSubCurrentRecord)) Else lngLastEntryUserID = 0
			If isDate(sarrySubForums(7,intSubCurrentRecord)) Then dtmLastEntryDate = CDate(sarrySubForums(7,intSubCurrentRecord)) Else dtmLastEntryDate = CDate("2001-01-01 00:00:00")
			If isNull(sarrySubForums(8, intSubCurrentRecord)) Then strForumPassword = "" Else strForumPassword = sarrySubForums(8, intSubCurrentRecord)
			blnForumLocked = CBool(sarrySubForums(9,intSubCurrentRecord))
			blnHideForum = CBool(sarrySubForums(10,intSubCurrentRecord))
			blnRead = CBool(sarrySubForums(11,intSubCurrentRecord))
			If isNumeric(sarrySubForums(12,intSubCurrentRecord)) Then lngTopicID = CLng(sarrySubForums(12,intSubCurrentRecord)) Else lngTopicID = 0
			strSubject = sarrySubForums(13,intSubCurrentRecord)	
			strForumImageIcon = sarrySubForums(14,intSubCurrentRecord)	

			'If this forum is to be hidden but the user is allowed access to it set the hidden boolen back to false
			If blnHideForum AND blnRead Then blnHideForum = False
				
			
			'Unread Posts *********
			intUnReadPostCount = 0
					
			'If there is a newer post than the last time the unread posts array was initilised run it again
			If dtmLastEntryDate > CDate(Session("dtmUnReadPostCheck")) Then Call UnreadPosts()
						
			'Count the number of unread posts in this forum
			If isArray(sarryUnReadPosts) AND dtmLastEntryDate > dtmLastVisitDate AND blnRead Then
				For intUnReadForumPostsLoop = 0 to UBound(sarryUnReadPosts,2)
					'Increament unread post count
					If CInt(sarryUnReadPosts(2,intUnReadForumPostsLoop)) = intSubForumID AND sarryUnReadPosts(3,intUnReadForumPostsLoop) = "1" Then intUnReadPostCount = intUnReadPostCount + 1
				Next
				
				'Get the text for unread post
				If intUnReadPostCount = 1 Then strNewPostText = strTxtNewPost Else strNewPostText = strTxtNewPosts	
			End If


			'If the forum is to be hidden then don't show it
			If blnHideForum = False Then

				'Get the row number
				intForumColourNumber = intForumColourNumber + 1
				
				'If mobile browser display different content
				If blnMobileBrowser Then
						
					'Calculate row colour
					Response.Write(vbCrLf & " <tr ")
					If (intForumColourNumber MOD 2 = 0 ) Then Response.Write("class=""evenTableRow"">") Else Response.Write("class=""oddTableRow"">") 
						
					'Display link to forum
					Response.Write("<td>")
					If blnUrlRewrite Then
						Response.Write("<a href=""" & SeoUrlTitle(strForumName, "") & "_forum" & intSubForumID & ".html" & strQsSID1 & """>" & strForumName & "</a>")
					Else
						Response.Write("<a href=""forum_topics.asp?FID=" & intSubForumID & strQsSID2 & SeoUrlTitle(strForumName, "&title=") & """>" & strForumName & "</a>")
					End If
					
					'Display unread post count to mobile users
					If intUnReadPostCount = 1 Then
						Response.Write(" [1 " & strTxtNewPost & "]")
					ElseIf intUnReadPostCount > 1 Then
						Response.Write(" [" & intUnReadPostCount & " " & strTxtNewPosts & "]")
					End If
					
					Response.Write("</td></tr>")	
					
					
					
				'Else not a mobile browser so display normal tables
				Else

					'Write the HTML of the forum descriptions and hyperlinks to the forums
	
					'Calculate row colour
					Response.Write(vbCrLf & " <tr ")
					If (intForumColourNumber MOD 2 = 0 ) Then Response.Write("class=""evenTableRow"">") Else Response.Write("class=""oddTableRow"">")
	
	
					'Display the status forum icons
					Response.Write(vbCrLf & "   <td align=""center"">")
					%><!-- #include file="includes/forum_status_icons_inc.asp" --><%
	     				Response.Write("</td>" & _
					vbCrLf & "  <td>")
	
	
					'Display forum
					If blnUrlRewrite Then
						Response.Write("<a href=""" & SeoUrlTitle(strForumName, "") & "_forum" & intSubForumID & ".html" & strQsSID1 & """>" & strForumName & "</a>")
					Else
						Response.Write("<a href=""forum_topics.asp?FID=" & intSubForumID & strQsSID2 & SeoUrlTitle(strForumName, "&title=") & """>" & strForumName & "</a>")
					End If
					
					
					'Display the number of people viewing in that forum
					If blnForumViewing AND blnActiveUsers Then 
						If viewingForum(intSubForumID) > 0 Then Response.Write(" <span class=""smText"">(" & viewingForum(intSubForumID) & " " & strTxtViewing & ")</span>")
					End If
					
					'Display forum details
					Response.Write("<br />" & strForumDiscription & "</td>" & _
					vbCrLf & "  <td align=""center"">" & lngNumberOfTopics & "</td>" & _
					vbCrLf & "  <td align=""center"">" & lngNumberOfPosts & "</td>" & _
					vbCrLf & "  <td class=""smText""  nowrap=""nowrap"">")
					If lngNumberOfPosts <> 0 Then 'Don't disply last post details if there are none
						
						'Don't dispaly details if the user has no read access on the forum
						If blnRead AND strForumPassword = "" Then
							
							'Display last post subject
							If blnUrlRewrite Then
								Response.Write("<a href=""" & SeoUrlTitle(strSubject, "") & "_topic" & lngTopicID & ".html" & strQsSID1 & """ title=""" & strSubject & """>")
							Else
								Response.Write("<a href=""forum_posts.asp?TID=" & lngTopicID & strQsSID2 & SeoUrlTitle(strSubject, "&title=") & """ title=""" & strSubject & """>")
							End If
							'Display Shorten subject (decode string first incase it contains allot of encoded characters)
							strSubject = TrimString(decodeString(strSubject), 30)
							strSubject = removeAllTags(strSubject)
							If blnBoldNewTopics AND intUnReadPostCount > 0 Then 'Unread topic subjects in bold
								Response.Write("<strong>" & strSubject & "</strong>")
							Else
								Response.Write(strSubject)
							End If
							Response.Write("</a><br />")
								
							'Who last post is by
							Response.Write(strTxtBy & "&nbsp;<a href=""member_profile.asp?PF=" & lngLastEntryUserID & strQsSID2 & """ class=""smLink"">" & strLastEntryUser & "</a> ")
							'If there are unread posts in the forum display differnt icon
							If intUnReadPostCount > 0 Then
								Response.Write("<a href=""get_last_post.asp?TID=" & lngTopicID & "&amp;ID=" & dblGetLastPostID & strQsSID2 & """><img src=""" & strImagePath & "view_unread_post." & strForumImageType & """ alt=""" & strTxtViewUnreadPost & " [" & intUnReadPostCount & " " & strNewPostText & "]"" title=""" & strTxtViewUnreadPost & " [" & intUnReadPostCount & " " & strTxtNewPosts & "]"" /></a>")
							
							'Else there are no unread posts so display a normal last post link
							Else
								Response.Write("<a href=""get_last_post.asp?TID=" & lngTopicID & "&amp;ID=" & dblGetLastPostID & strQsSID2 & """><img src=""" & strImagePath & "view_last_post." & strForumImageType & """ alt=""" & strTxtViewLastPost & """ title=""" & strTxtViewLastPost & """ /></a>")
							End If
							
						End If
						'Last post date
						Response.Write("<br />" & DateFormat(dtmLastEntryDate) & "&nbsp;" &  strTxtAt & "&nbsp;" & TimeFormat(dtmLastEntryDate))
					End If
					Response.Write("</td>"  & _
					vbCrLf & " </tr>")
				End If


			End If

			'Move to the next database record
			intSubCurrentRecord = intSubCurrentRecord + 1


			'If there are more records in the array to display then run some test to see what record to display next and where
			If intSubCurrentRecord <= Ubound(sarrySubForums,2) Then

				'Becuase the member may have an individual permission entry in the permissions table for this forum,
				'it maybe listed twice in the array, so we need to make sure we don't display the same forum twice
				If intSubForumID = CInt(sarrySubForums(0,intSubCurrentRecord)) Then intSubCurrentRecord = intSubCurrentRecord + 1

				'If there are no records left exit loop
				If intSubCurrentRecord > Ubound(sarrySubForums,2) Then Exit Do
			End If

		'Loop back round for next forum
		Loop

%>
</table>
<br /><%
	End If
End If





'******************************************
'***	      Display Topics		***
'******************************************

%>
<table class="basicTable" cellspacing="1" cellpadding="3" align="center">
 <tr>
  <td width="<% If blnPollCreate Then Response.Write("205") Else Response.Write("100") %>">
   <a href="new_topic_form.asp?FID=<% = intForumID & strQsSID2 %>" title="<% = strTxtCreateNewTopic %>" class="largeButton" rel="nofollow">&nbsp;<% = strTxtNewTopic %> <img src="<% = strImagePath %>new_topic.<% = strForumImageType %>" alt="<% = strTxtCreateNewTopic %>" /></a><%

'If the user can create a poll disply a create poll link
If blnPollCreate Then Response.Write ("<a href=""new_poll_form.asp?FID=" & intForumID & strQsSID2 & """ title=""" & strTxtCreateNewPoll & """ class=""largeButton"">&nbsp;&nbsp;" & strTxtNewPoll & "&nbsp;&nbsp;<img src=""" & strImagePath & "new_poll." & strForumImageType & """ alt=""" & strTxtCreateNewPoll & """></a>")

'Do not show any drop down for mobile browsers
If blnMobileBrowser = False Then
%>
  </td>
  <td><% = strTxtShowTopics %>
   <select name="show" id="show" onchange="ShowTopics(this)">
    <option value="0"<% If intShowTopicsFrom = 0 Then Response.Write " selected=""selected""" %>><% = strTxtAnyDate %></option>
    <option value="1"<% If intShowTopicsFrom = 1 Then Response.Write " selected=""selected""" %>><% = DateFormat(dtmLastVisitDate) & " " & strTxtAt & " " & TimeFormat(dtmLastVisitDate) %></option>
    <option value="2"<% If intShowTopicsFrom = 2 Then Response.Write " selected=""selected""" %>><% = strTxtYesterday %></option>
    <option value="3"<% If intShowTopicsFrom = 3 Then Response.Write " selected=""selected""" %>><% = strTxtLastTwoDays %></option>
    <option value="4"<% If intShowTopicsFrom = 4 Then Response.Write " selected=""selected""" %>><% = strTxtLastWeek %></option>
    <option value="5"<% If intShowTopicsFrom = 5 Then Response.Write " selected=""selected""" %>><% = strTxtLastMonth %></option>
    <option value="6"<% If intShowTopicsFrom = 6 Then Response.Write " selected=""selected""" %>><% = strTxtLastTwoMonths %></option>
    <option value="7"<% If intShowTopicsFrom = 7 Then Response.Write " selected=""selected""" %>><% = strTxtLastSixMonths %></option>
    <option value="8"<% If intShowTopicsFrom = 8 Then Response.Write " selected=""selected""" %>><% = strTxtLastYear %></option>
   </select>
  </td><%
  
End If

%>
  <td align="right" nowrap="nowrap"><!-- #include file="includes/page_link_inc.asp" --></td>
 </tr>
</table>
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center"><%

'Display column headings if not a mobile
If blnMobileBrowser = False Then

%>
 <tr class="tableLedger">
   <td width="5%">&nbsp;</td>
   <td width="50%"><div style="float:left;"><a href="forum_topics.asp?FID=<% = intForumID %>&amp;SO=T<% = strQsSID2 %>" title="<% = strTxtReverseSortOrder %>"><% = strTxtTopics %></a><% If UCase(Request.QueryString("SO")) = "T" Then Response.Write(" <a href=""forum_topics.asp?FID=" & intForumID & "&amp;SO=T&amp;OB=" & strSortDirection & strQsSID2 & """><img src=""" & strImagePath & strSortDirection & "." & strForumImageType & """ title=""" & strTxtReverseSortOrder & """ alt=""" & strTxtReverseSortOrder & """ /></a>") %> / <a href="forum_topics.asp?FID=<% = intForumID %>&amp;SO=A<% = strQsSID2 %>" title="<% = strTxtReverseSortOrder %>"><% = strTxtThreadStarter %></a><% If UCase(Request.QueryString("SO")) = "A" Then Response.Write(" <a href=""forum_topics.asp?FID=" & intForumID & "&amp;SO=A&amp;OB=" & strSortDirection & strQsSID2 & """><img src=""" & strImagePath & strSortDirection & "." & strForumImageType & """ alt=""" & strTxtReverseSortOrder & """ title=""" & strTxtReverseSortOrder & """ /></a>") %></div><% 
   	'If rating is enabled
   	If blnTopicRating Then 
   		%><div style="float:right;"><a href="forum_topics.asp?FID=<% = intForumID %>&amp;SO=UR<% = strQsSID2 %>&amp;OB=desc" title="<% = strTxtReverseSortOrder %>"><% = strTxtRating %></a><% If UCase(Request.QueryString("SO")) = "UR" Then Response.Write(" <a href=""forum_topics.asp?FID=" & intForumID & "&amp;SO=UR&amp;OB=" & strSortDirection & strQsSID2 & """><img src=""" & strImagePath & strSortDirection & "." & strForumImageType & """ alt=""" & strTxtReverseSortOrder & """ title=""" & strTxtReverseSortOrder & """ /></a>") %>&nbsp;</div><%
   		
   	End If 
   %></td>
   <td width="10%" align="center" nowrap="nowrap"><a href="forum_topics.asp?FID=<% = intForumID %>&amp;SO=R<% = strQsSID2 %>" title="<% = strTxtReverseSortOrder %>"><% = strTxtReplies %></a><% If UCase(Request.QueryString("SO")) = "R" Then Response.Write(" <a href=""forum_topics.asp?FID=" & intForumID & "&amp;SO=R&amp;OB=" & strSortDirection & strQsSID2 & """><img src=""" & strImagePath & strSortDirection & "." & strForumImageType & """ title=""" & strTxtReverseSortOrder & """ alt=""" & strTxtReverseSortOrder & """ /></a>") %></td>
   <td width="10%" align="center" nowrap="nowrap"><a href="forum_topics.asp?FID=<% = intForumID %>&amp;SO=V<% = strQsSID2 %>" title="<% = strTxtReverseSortOrder %>"><% = strTxtViews %></a><% If UCase(Request.QueryString("SO")) = "V" Then Response.Write(" <a href=""forum_topics.asp?FID=" & intForumID & "&amp;SO=V&amp;OB=" & strSortDirection & strQsSID2 & """><img src=""" & strImagePath & strSortDirection & "." & strForumImageType & """ title=""" & strTxtReverseSortOrder & """ alt=""" & strTxtReverseSortOrder & """ /></a>") %></td>
   <td width="30%"><a href="forum_topics.asp?FID=<% = intForumID %><% = strQsSID2 %>" title="<% = strTxtReverseSortOrder %>"><% = strTxtLastPost %></a><% If UCase(Request.QueryString("SO")) = "" Then Response.Write(" <a href=""forum_topics.asp?FID=" & intForumID & "&amp;OB=" & strSortDirection & strQsSID2 & """><img src=""" & strImagePath & strSortDirection & "." & strForumImageType & """ title=""" & strTxtReverseSortOrder & """ alt=""" & strTxtReverseSortOrder & """ /></a>") %></td>
  </tr><%
End If
  

'If there are no topics to display, show a message saying so
If lngTotalRecords <= 0 Then

	'If there are no Topic's to display then display the appropriate error message
	Response.Write vbCrLf & "  <tr class=""tableRow""><td colspan=""5"" align=""center""><br />" & strTxtNoTopicsToDisplay & "&nbsp;" & strShowTopicsFrom & "<br /><br /></td></tr>"

'Else there the are topic's so write the HTML to display the topic names and a discription
Else

	'Do....While Loop to loop through the recorset to display the forum topics
	Do While intCurrentRecord < intEndPosition
		
		

		'If there are no topic records left to display then exit loop
		If intCurrentRecord >= lngTotalRecords Then Exit Do

		'SQL Query Array Look Up table
		'0 = tblTopic.Topic_ID
		'1 = tblTopic.Poll_ID
		'2 = tblTopic.Moved_ID
		'3 = tblTopic.Subject
		'4 = tblTopic.Icon
		'5 = tblTopic.Start_Thread_ID
		'6 = tblTopic.Last_Thread_ID
		'7 = tblTopic.No_of_replies
		'8 = tblTopic.No_of_views
		'9 = tblTopic.Locked
		'10 = tblTopic.Priority
		'11 = tblTopic.Hide
		'12 = tblThread.Message_date
		'13 = tblThread.Message,
		'14 = tblThread.Author_ID,
		'15 = tblAuthor.Username,
		'16 = LastThread.Message_date,
		'17 = LastThread.Author_ID,
		'18 = LastAuthor.Username
		'19 = tblTopic.Event_date
		'20 = tblTopic.Event_date_end
		'21 = tblTopic.Rating
		'22 = tblTopic.Rating_Votes


		'Read in Topic details from the database
		lngTopicID = CLng(sarryTopics(0,intCurrentRecord))
		lngPollID = CLng(sarryTopics(1,intCurrentRecord))
		intMovedForumID = CInt(sarryTopics(2,intCurrentRecord))
		strSubject = sarryTopics(3,intCurrentRecord)
		strTopicIcon = sarryTopics(4,intCurrentRecord)
		lngNumberOfReplies = CLng(sarryTopics(7,intCurrentRecord))
		lngNumberOfViews = CLng(sarryTopics(8,intCurrentRecord))
		blnTopicLocked = CBool(sarryTopics(9,intCurrentRecord))
		intPriority = CInt(sarryTopics(10,intCurrentRecord))
		blnHideTopic = CBool(sarryTopics(11,intCurrentRecord))
		dtmEventDate = sarryTopics(19,intCurrentRecord)
		dtmEventDateEnd = sarryTopics(20,intCurrentRecord)
		If isNumeric(sarryTopics(21,intCurrentRecord)) Then dblTopicRating = CDbl(sarryTopics(21,intCurrentRecord)) Else dblTopicRating = 0
		If isNumeric(sarryTopics(22,intCurrentRecord)) Then lngTopicVotes = CLng(sarryTopics(22,intCurrentRecord)) Else lngTopicVotes = 0
	
		'Read in the first post details
		dtmFirstEntryDate = CDate(sarryTopics(12,intCurrentRecord))
		strFirstPostMsg = Mid(sarryTopics(13,intCurrentRecord), 1, 275)
		lngTopicStartUserID = CLng(sarryTopics(14,intCurrentRecord))
		strTopicStartUsername = sarryTopics(15,intCurrentRecord)

		'Read in the last post details
		lngLastEntryMessageID = CLng(sarryTopics(6,intCurrentRecord))
		dtmLastEntryDate = CDate(sarryTopics(16,intCurrentRecord))
		lngLastEntryUserID = CLng(sarryTopics(17,intCurrentRecord))
		strLastEntryUsername = sarryTopics(18,intCurrentRecord)
		

		'Clean up input to prevent XXS hack
		strSubject = formatInput(strSubject)

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
		

		'Get a record number for the number of non priority posts
		If intPriority <= 1 Then intNonPriorityTopicNum = intNonPriorityTopicNum + 1

		'If this is the first topic that is not important then display the forum topics bar
		If intNonPriorityTopicNum = 1 Then Response.Write vbCrLf & "    <tr class=""tableSubLedger""><td colspan=""5"">" & strTxtForum & " " & strTxtTopics & "</td></tr>"

		'If this is the first topic and it is an important one then display a bar saying so
		If intCurrentRecord = 0 AND intPriority => 2 Then Response.Write vbCrLf & "<tr class=""tableSubLedger""><td colspan=""5"">" & strTxtAnnouncements & "</td></tr>"


		


		'Calculate the row colour
		If intCurrentRecord MOD 2=0 Then strTableRowColour = "evenTableRow" Else strTableRowColour = "oddTableRow"

		'If this is a hidden post then change the row colour to highlight it
		If blnHideTopic Then strTableRowColour = "hiddenTableRow"
			
			
		'If mobile browser display different content
		If blnMobileBrowser Then
			
			'Write the HTML of the Topic descriptions as hyperlinks to the Topic details and message
			Response.Write(vbCrLf & "  <tr class=""" & strTableRowColour & """>")
			
			'Display the subject of the topic
			Response.Write("<td>")
			'If sticky prefix with sticky
			If  intPriority = 1 Then
				Response.Write(strTxtSticky & ": ")
			End If
			'If poll prefix with Poll
			If lngPollID > 0 Then 
				Response.Write(strTxtPoll2 & ": ")
			End If
			'If event prefix with Event
			If isDate(dtmEventDate) Then
				Response.Write(strTxtEvent & ": ")
			End If
			If blnUrlRewrite Then
				Response.Write("<a href=""" & SeoUrlTitle(strSubject, "") & "_topic" & lngTopicID & ".html")
				If intPriority = 3 Then Response.Write("?FID=" & intForumID & "&amp;PR=3" & strQsSID2) Else Response.Write(strQsSID1)
				Response.Write(""">" & strSubject & "</a>")
			Else
				Response.Write("<a href=""forum_posts.asp?TID=" & lngTopicID)
				If intPriority = 3 Then Response.Write("&amp;FID=" & intForumID & "&amp;PR=3")
				Response.Write("" & strQsSID2 & SeoUrlTitle(strSubject, "&title=") & """>" & strSubject & "</a>")
			End If
			
			'Display the number of unread posts
			If intUnReadPostCount > 0 Then
		   		Response.Write(" [<a href=""get_last_post.asp?TID=" & lngTopicID & "&amp;ID=" & dblGetLastPostID)
		   		If intPriority = 3 Then Response.Write("&amp;FID=" & intForumID & "&amp;PR=3") 
				Response.Write("" & strQsSID2 & """ title=""" & strTxtViewUnreadPost & """>" & intUnReadPostCount & " " & strNewPostText & "</a>]") 
			End If
			
			'Display last post details
			Response.Write("<br /><span class=""smText"">" & strTxtLastPost & ": " & strLastEntryUsername  & " " &  DateFormat(dtmLastEntryDate) & " " & strTxtAt & " " & TimeFormat(dtmLastEntryDate) & "</span>")
		
			Response.Write("</td></tr>")	


		'Else table view for all other browser
		Else


			'Write the HTML of the Topic descriptions as hyperlinks to the Topic details and message
			Response.Write(vbCrLf & "  <tr class=""" & strTableRowColour & """>")
	
			
			'Display the status topic icons
			Response.Write(vbCrLf & "   <td align=""center"">")
			%><!-- #include file="includes/topic_status_icons_inc.asp" --><%
	     		Response.Write("</td>")
	
	
	
	
	     		Response.Write(vbCrLf & "   <td><div style=""float:left"">")
	     	
			 'If the user is a forum admin or a moderator then give let them delete the topic
			 If blnAdmin  OR blnModerator Then 
			 	
			 	Response.Write("<span id=""modTools" & lngTopicID & """ onclick=""showDropDown('modTools" & lngTopicID & "', 'modToolsMenu" & lngTopicID & "', 120, 0);"" class=""dropDownPointer""><img src=""" & strImagePath & "moderator_tools." & strForumImageType & """ alt=""" & strTxtModeratorTools & """ title=""" & strTxtModeratorTools & """ /></span> " & _
				"<div id=""modToolsMenu" & lngTopicID & """ class=""dropDownStatusBar"">" & _
				"<a href=""javascript:winOpener('pop_up_topic_admin.asp?TID=" & lngTopicID & strQsSID2 & "','admin',1,1,600,285)""><div>" & strTxtTopicAdmin & "</div></a>")
				
				'Lock or un-lock forum if admin
				If blnTopicLocked Then
					Response.Write("<a href=""lock_topic.asp?mode=UnLock&amp;TID=" & lngTopicID & "&amp;FID=" & intForumID & "&amp;XID=" & getSessionItem("KEY") & strQsSID2 & """><div>" & strTxtUnLockTopic & "</div></a>")
				Else
					Response.Write("<a href=""lock_topic.asp?mode=Lock&amp;TID=" & lngTopicID & "&amp;FID=" & intForumID & "&amp;XID=" & getSessionItem("KEY") & strQsSID2 & """><div>" & strTxtLockTopic & "</div></a>")
				End If
	
				'Hide or show topic
				If blnHideTopic = false Then
					Response.Write("<a href=""lock_topic.asp?mode=Hide&amp;TID=" & lngTopicID & "&amp;FID=" & intForumID & "&amp;XID=" & getSessionItem("KEY") & strQsSID2 & """><div>" & strTxtHideTopic & "</div></a>")
				Else
					Response.Write("<a href=""lock_topic.asp?mode=Show&amp;TID=" & lngTopicID & "&amp;FID=" & intForumID & "&amp;XID=" & getSessionItem("KEY") & strQsSID2 & """><div>" & strTxtShowTopic & "</div></a>")
				End If
				
				Response.Write("<a href=""delete_topic.asp?TID=" & lngTopicID & "&amp;PN=" & intRecordPositionPageNum & "&amp;XID=" & getSessionItem("KEY") & strQsSID2 & """ onclick=""return confirm('" & strTxtDeleteTopicAlert & "')""><div>" & strTxtDeleteTopic & "</div></a>")
				Response.Write("</div>")
			 	
			End If
			
			
			
			'If topic icons enabled and we have a topic icon display it	
	     		If blnTopicIcon AND strTopicIcon <> "" Then Response.Write("<img src=""" & strTopicIcon & """ alt=""" & strTxtMessageIcon & """ title=""" & strTxtMessageIcon & """ /> ")
	
	
	
			'Display the subject of the topic
			If blnUrlRewrite Then
				Response.Write("<a href=""" & SeoUrlTitle(strSubject, "") & "_topic" & lngTopicID & ".html")
				If intPriority = 3 Then Response.Write("?FID=" & intForumID & "&amp;PR=3" & strQsSID2) Else Response.Write(strQsSID1)
				Response.Write(""" title=""" & strFirstPostMsg & """>")
				
			Else
				Response.Write("<a href=""forum_posts.asp?TID=" & lngTopicID)
				If intPriority = 3 Then Response.Write("&amp;FID=" & intForumID & "&amp;PR=3")
				Response.Write("" & strQsSID2 & SeoUrlTitle(strSubject, "&title=") & """ title=""" & strFirstPostMsg & """>")
			End If
				
			If blnBoldNewTopics AND intUnReadPostCount > 0 Then 'Unread topic subjects in bold
				Response.Write("<strong>" & strSubject & "</strong></a>")
				'Display the number of unread posts
				If intUnReadPostCount > 0 Then
			   		Response.Write(" <span class=""smText"">(<a href=""get_last_post.asp?TID=" & lngTopicID & "&amp;ID=" & dblGetLastPostID)
			   		If intPriority = 3 Then Response.Write("&amp;FID=" & intForumID & "&amp;PR=3") 
					Response.Write("" & strQsSID2 & """ title=""" & strTxtViewUnreadPost & """ class=""smLink"">" & intUnReadPostCount & " " & strNewPostText & "</a>)</span>") 
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
			 				If blnUrlRewrite Then
								Response.Write("<a href=""" & SeoUrlTitle(strSubject, "") & "_topic" & lngTopicID & "_page4.html")
								If intPriority = 3 Then Response.Write("?FID=" & intForumID & "&amp;PR=3" & strQsSID2) Else Response.Write(strQsSID1)
								Response.Write(""" class=""smPageLink"" title=""" & strTxtPage & " 4"">4</a>")
								
							Else
								Response.Write("<a href=""forum_posts.asp?TID=" & lngTopicID & "&amp;PN=4")
								'If a priority topic need to make sure we don't change forum
					 			If intPriority = 3 Then Response.Write("&amp;FID=" & intForumID & "&amp;PR=3")
					 			Response.Write("" & strQsSID2 & SeoUrlTitle(strSubject, "&amp;title=") & """ class=""smPageLink"" title=""" & strTxtPage & " 4"">4</a>")
							End If
			 	
						'Else display the last 2 pages
						Else
	
							Response.Write("&nbsp;")
							
							'2nd form last page
							If blnUrlRewrite Then
								Response.Write("<a href=""" & SeoUrlTitle(strSubject, "") & "_topic" & lngTopicID & "_page" & intNumberOfTopicPages - 1 & ".html")
								If intPriority = 3 Then Response.Write("?FID=" & intForumID & "&amp;PR=3" & strQsSID2) Else Response.Write(strQsSID1)
								Response.Write(""" class=""smPageLink"" title=""" & strTxtPage & " " & intNumberOfTopicPages - 1 & """>" & intNumberOfTopicPages - 1 & "</a>")
								
							Else
								Response.Write("<a href=""forum_posts.asp?TID=" & lngTopicID & "&amp;PN=" & intNumberOfTopicPages - 1)
								'If a priority topic need to make sure we don't change forum
				 				If intPriority = 3 Then Response.Write("&amp;FID=" & intForumID & "&amp;PR=3")
				 				Response.Write("" & strQsSID2 & SeoUrlTitle(strSubject, "&amp;title=") & """ class=""smPageLink"" title=""" & strTxtPage & " " & intNumberOfTopicPages - 1 & """>" & intNumberOfTopicPages - 1 & "</a>")
							End If
							
							'Last page
							If blnUrlRewrite Then
								Response.Write("<a href=""" & SeoUrlTitle(strSubject, "") & "_topic" & lngTopicID & "_page" & intNumberOfTopicPages & ".html")
								If intPriority = 3 Then Response.Write("?FID=" & intForumID & "&amp;PR=3" & strQsSID2) Else Response.Write(strQsSID1)
								Response.Write(""" class=""smPageLink"" title=""" & strTxtPage & " " & intNumberOfTopicPages & """>" & intNumberOfTopicPages & "</a>")
								
							Else
								Response.Write("<a href=""forum_posts.asp?TID=" & lngTopicID & "&amp;PN=" & intNumberOfTopicPages)
								'If a priority topic need to make sure we don't change forum
				 				If intPriority = 3 Then Response.Write("&amp;FID=" & intForumID & "&amp;PR=3")
				 				Response.Write("" & strQsSID2 & SeoUrlTitle(strSubject, "&amp;title=") & """ class=""smPageLink"" title=""" & strTxtPage & " " & intNumberOfTopicPages & """>" & intNumberOfTopicPages & "</a>")
							End If
						End If
	
						'Exit the loop as we are finshed here
			 			Exit For
			 		End If
			 		
			 		'If page 1 don't include page number
			 		If intTopicPagesLoopCounter = 1 Then
			 			'Display the links to the other pages
				 		If blnUrlRewrite Then
							Response.Write("<a href=""" & SeoUrlTitle(strSubject, "") & "_topic" & lngTopicID & ".html")
							'If a priority topic need to make sure we don't change forum
							If intPriority = 3 Then Response.Write("?FID=" & intForumID & "&amp;PR=3" & strQsSID2) Else Response.Write(strQsSID1)
							Response.Write(""" class=""smPageLink"" title=""" & strTxtPage & " 1"">1</a>")
							
						Else
							Response.Write("<a href=""forum_posts.asp?TID=" & lngTopicID)
				 			'If a priority topic need to make sure we don't change forum
				 			If intPriority = 3 Then Response.Write("&amp;FID=" & intForumID & "&amp;PR=3")
				 			Response.Write("" & strQsSID2 & SeoUrlTitle(strSubject, "&amp;title=") & """ class=""smPageLink"" title=""" & strTxtPage & " 1"">1</a>")
						End If
					
					Else
			 		
				 		'Display the links to the other pages
				 		If blnUrlRewrite Then
							Response.Write("<a href=""" & SeoUrlTitle(strSubject, "") & "_topic" & lngTopicID & "_page" & intTopicPagesLoopCounter & ".html")
							'If a priority topic need to make sure we don't change forum
							If intPriority = 3 Then Response.Write("?FID=" & intForumID & "&amp;PR=3" & strQsSID2) Else Response.Write(strQsSID1)
							Response.Write(""" class=""smPageLink"" title=""" & strTxtPage & " " & intTopicPagesLoopCounter & """>" & intTopicPagesLoopCounter & "</a>")
							
						Else
							Response.Write("<a href=""forum_posts.asp?TID=" & lngTopicID & "&amp;PN=" & intTopicPagesLoopCounter)
				 			'If a priority topic need to make sure we don't change forum
				 			If intPriority = 3 Then Response.Write("&amp;FID=" & intForumID & "&amp;PR=3")
				 			Response.Write("" & strQsSID2 & SeoUrlTitle(strSubject, "&amp;title=") & """ class=""smPageLink"" title=""" & strTxtPage & " " & intTopicPagesLoopCounter & """>" & intTopicPagesLoopCounter & "</a>")
						End If
					End If
		
			 		
			 	Next
			 	Response.Write("</div>")
			 End If
		 
		
		  %></td>
   <td align="center"><% = lngNumberOfReplies %></td>
   <td align="center"><% = lngNumberOfViews %></td>
   <td class="smText" nowrap="nowrap"><% = strTxtBy %> <a href="member_profile.asp?PF=<% = lngLastEntryUserID & strQsSID2 %>"  class="smLink" rel="nofollow"><% = strLastEntryUsername %></a><br/> <% = DateFormat(dtmLastEntryDate) & " " & strTxtAt & " " & TimeFormat(dtmLastEntryDate) %> <%
   
	   		'If there are unread posts display a differnet icon and link to the last unread post
	   		If intUnReadPostCount > 0 Then
	   			 
	   			Response.Write("<a href=""get_last_post.asp?TID=" & lngTopicID & "&amp;ID=" & dblGetLastPostID)
	   			If intPriority = 3 Then Response.Write("&amp;FID=" & intForumID & "&amp;PR=3") 
				Response.Write("" & strQsSID2 & """><img src=""" & strImagePath & "view_unread_post." & strForumImageType & """ alt=""" & strTxtViewUnreadPost & " [" & intUnReadPostCount & " " & strNewPostText & "]"" title=""" & strTxtViewUnreadPost & " [" & intUnReadPostCount & " " & strNewPostText & "]"" /></a> ") 
	   	
	   		'Else there are no unread posts so display a normal topic link
	   		Else
				Response.Write("<a href=""get_last_post.asp?TID=" & lngTopicID & "&amp;ID=" & dblGetLastPostID)
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
  <td>
   <a href="new_topic_form.asp?FID=<% = intForumID & strQsSID2 %>" title="<% = strTxtCreateNewTopic %>" class="largeButton" rel="nofollow">&nbsp;<% = strTxtNewTopic %> <img src="<% = strImagePath %>new_topic.<% = strForumImageType %>" alt="<% = strTxtCreateNewTopic %>" /></a><%

'If the user can create a poll disply a create poll link
If blnPollCreate Then Response.Write ("<a href=""new_poll_form.asp?FID=" & intForumID & strQsSID2 & """ title=""" & strTxtCreateNewPoll & """ class=""largeButton"">&nbsp;&nbsp;" & strTxtNewPoll & "&nbsp;&nbsp;<img src=""" & strImagePath & "new_poll." & strForumImageType & """  alt=""" & strTxtCreateNewPoll & """></a>")

%>
  </td>
  <td align="right" nowrap="nowrap"><!-- #include file="includes/page_link_inc.asp" --></td>
 </tr>
</table>
<table class="basicTable" cellspacing="0" cellpadding="4" align="center">
 <tr>
  <td valign="top"><br /><!-- #include file="includes/forum_jump_inc.asp" --></td>
  <td class="smText" align="right" nowrap="nowrap"><!-- #include file="includes/show_forum_permissions_inc.asp" --></td>
 </tr>
</table>
<br />
<div align="center">
<%


'Reset Server Objects
Call closeDatabase()

'If a mobile browser display an option to switch to and from mobile view
If blnMobileBrowser Then 
	Response.Write (strTxtViewIn & ": <strong>" & strTxtMoble & "</strong> | <a href=""" & strLinkPage & "MobileView=off" & strQsSID2 & """ rel=""nofollow"">" & strTxtClassic & "</a><br /><br />")
ElseIf blnMobileClassicView Then
	Response.Write (strTxtViewIn & ": <a href=""" & strLinkPage & "MobileView=on" & strQsSID2 & """ rel=""nofollow"">" & strTxtMoble & "</a> | <strong>" & strTxtClassic & "</strong><br /><br />")
End If


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
</div><%

'Display an alert message letting the user know the topic has been deleted
If Request.QueryString("DL") = "1" Then
	Response.Write("<script language=""JavaScript"" type=""text/javascript"">")
	Response.Write("alert('" & strTxtTheTopicIsNowDeleted & "')")
	Response.Write("</script>")
End If

'Display an alert message if the user is watching this forum for email notification
If Request.QueryString("EN") = "FS" Then
	Response.Write("<script language=""JavaScript"" type=""text/javascript"">")
	Response.Write("alert('" & strTxtYouAreNowBeNotifiedOfPostsInThisForum & "')")
	Response.Write("</script>")
End If

'Display an alert message if the user is not watching this forum for email notification
If Request.QueryString("EN") = "FU" Then
	Response.Write("<script language=""JavaScript"" type=""text/javascript"">")
	Response.Write("alert('" & strTxtYouAreNowNOTBeNotifiedOfPostsInThisForum & "')")
	Response.Write("</script>")
End If

%>
<!-- #include file="includes/footer.asp" -->