<% @ Language=VBScript %>
<% Option Explicit %>
<% Response.AddHeader "X-UA-Compatible","IE=EmulateIE10" %>
<!--#include file="common.asp" -->
<!--#include file="functions/functions_date_time_format.asp" -->
<!--#include file="functions/functions_format_post.asp" -->
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






Response.Buffer = True

'Make sure this page is not cached
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"




'Dimension variables
Dim sarryPosts			'Holds the posts recordset
Dim strForumName		'Holds the forum name
Dim intCatID			'Holds the cat ID
Dim strCatName			'Holds the cat name
Dim intMasterForumID		'Holds the main forum ID
Dim strMasterForumName		'Holds the main forum name
Dim lngTopicID			'Holds the topic number
Dim strSubject			'Holds the topic subject
Dim strUsername 		'Holds the Username of the thread
Dim strAuthorEmail
Dim lngUserID			'Holds the ID number of the user
Dim dtmPostDate			'Holds the date the thread was made
Dim strMessage			'Holds the post to display
Dim strTopicIcon		'Holds the topic icon
Dim lngMessageID		'Holds the message ID number
Dim strAuthorHomepage		'Holds the homepage of the Username if it is given
Dim strAuthorLocation		'Holds the location of the user if given
Dim strAuthorAvatar		'Holds the authors avatar
Dim strAuthorSignature		'Holds the authors signature
Dim lngAuthorNumOfPosts		'Holds the number of posts the user has made to the forum
Dim dtmAuthorRegistration	'Holds the registration date of the user
Dim lngNumberOfViews		'Holds the number of times the topic has been viewed to save back to the database
Dim intStatus			'Holds the users interger status
Dim strStatus			'Holds the users status
Dim strMode			'Holds the mode of the page that is being passed
Dim intTopicPageNumber		'Holds the topic page position to link back to
Dim blnIsModerator		'Set to true if the user who posted the message is a forum moderator
Dim blnForumLocked		'Set to true if the forum is locked
Dim blnTopicLocked		'set to true if the topic is locked
Dim intThreadNo			'Holds the number of threads in the topic
Dim intPriority			'Holds the priority level of the topic
Dim intRecordPositionPageNum	'Holds the recorset page number to show the Threads for
Dim lngTotalRecordsPages	'Holds the number of pages
Dim intTopicPageLoopCounter	'Loop counter for other thread page link
Dim lngTotalRecords		'Holds the total number of therads in this topic
Dim strAuthorIP			'Holds the authors IP
Dim strSearchKeywords		'Holds the keywords to search for
Dim sarySearchWord		'Array to hold the search words
Dim strGroupName		'Holds the authors group name
Dim intRankStars		'Holds the number of stars for the group
Dim strRankCustomStars		'Holds custom stars for the user group
Dim lngPollID			'Holds the poll ID
Dim blnPollNoReply		'Set to true if users can't reply to a poll
Dim blnBannedIP			'Set to true if the user is using a banned IP
Dim intIndexPosition		'Holds the idex poistion in the emiticon array
Dim intNumberOfOuterLoops	'Holds the outer loop number for rows
Dim intLoop			'Holds the loop index position
Dim intInnerLoop		'Holds the inner loop number for columns
Dim strMemberTitle		'Holds the members title
Dim blnIsUserOnline		'Set to true if the user is online
Dim blnHidePost			'Set to true if post is hidden
Dim strPostTableRowClass	'Post row colour
Dim intStartPosition		'Holds the start poition for records to be shown
Dim intEndPosition		'Holds the end poition for records to be shown
Dim intCurrentRecord		'Holds the current records for the posts
Dim strGuestUsername		'Holds the Guest Username if it is a guest posting
Dim blnHideTopic		'Set to true if topic is hidden
Dim lngGetPostID		'Holds the post to go straight to
Dim strSortBy			'Holds the direction the posts are listed in
Dim strPostTableSideClass	'Holds the post message side column css class
Dim strPostTableTopClass	'Holds the post table top details
Dim dtmEventDate		'Holds the date if this is a calendar event
Dim dtmEventDateEnd		'Holds the date if this is a calendar event
Dim intPageLinkLoopCounter	'Loop counter for page links include
Dim intPostRowsSpan		'Holds the rowspam number for post table formatting
Dim strForumPassword		'Holds the forum password
Dim intUnReadForumPostsLoop	'Loop counter for unread loop array
Dim intAutoLockTopicsAmount	'Holds the amount topics are auto closed at
Dim blnUnreadPost		'Set to true if the post is un-read
Dim strMetaDescription		'Used for meta tags
Dim lngAuthorNumOfPoints	'Number of points for each user
Dim strSQLFromWhere		'Used for the From where clause
Dim lngPrePostCount		'Used to count the number of posts before the post we want to display
Dim dblTopicRating		'Holds the rating for a topic
Dim lngTopicVotes		'Number of votes a topic receives
Dim strDynamicKeywords		'Holds the keywords for the page
Dim strPageQueryString		'Holds the querystring for the page	
Dim intUnReadPostCount		'Holds the unread post count
Dim lngFirstUnreadPostID	'Holds the ID of the first unread post in the topic
Dim blnSuspended		'Set to true if member suspended
Dim strFormID			'Holds the form ID value
Dim blnShowSignature		'Set to true if showing signature for this group
Dim strCanonicalURL
Dim lngPostThanks		'Holds the number of thanks a topic has received
Dim blnTopicAnswer		'Holds if the topic is an answer
Dim saryAnswerPosts		'Holds the anwser posts
Dim intCurrentAnswerRecord
Dim blnTopTopicAnswer		'Set to true if an anwser post display at the top
Dim intTopTopicAnswers		'Holds the number of topic anwsers
Dim strSocialNetButtonsURL


'Initialise variables
strMode = "reply"
lngMessageID = 0
lngGetPostID = 0
intForumID = 0
lngTopicID = 0
intThreadNo = 0
blnIsModerator = False
blnPollNoReply = False
blnBannedIP = False
blnHidePost = False
intPriority = 0
blnHideTopic = False
blnUnreadPost = False
intTopTopicAnswers = -1
intCurrentAnswerRecord = 0



'Read in the querystrings for the page
strPageQueryString = Request.QueryString()




'Remove the page title from the querystring beofre doing the sql injection test
If Request.QueryString("title") <> "" Then strPageQueryString = Replace(strPageQueryString, Request.QueryString("title"), "")

'Read in the keywords if comming from a search
strSearchKeywords = Trim(Mid(Request.QueryString("KW"), 1, 35))
strSearchKeywords = Replace(strSearchKeywords, ";", " ")
'Split up the keywords to be searched
sarySearchWord = Split(Trim(strSearchKeywords), " ")

'Decode URL underscore and dot to prevent keyword search issues
strPageQueryString = Replace(strPageQueryString, "%5F", "_")
strPageQueryString = Replace(strPageQueryString, "%2E", ".")


'Remove the keyword from the querystring before sql injection test
For intLoop = 0 to UBound(sarySearchWord)
	strPageQueryString = Replace(strPageQueryString, sarySearchWord(intLoop), "")
Next


'Test querystrings for any SQL Injection keywords
Call SqlInjectionTest(strPageQueryString)




'See if the user is using a banned IP address
If bannedIP() Then
	'If the user is using a banned IP then set the banned IP variable to true and active member variable to false
	blnBannedIP = True
End If


'If this is the first time the page is displayed then the Forum Thread record position is set to page 1
If isNumeric(Request.QueryString("PN")) = false Then
	intRecordPositionPageNum = 1
ElseIf Request.QueryString("PN") < 1 Then
	intRecordPositionPageNum = 1
'Else the page has been displayed before so the Forum Thread record postion is set to the Record Position number
Else
	intRecordPositionPageNum = IntC(Request.QueryString("PN"))
End If



'Read in the Topic ID for the topic to display
If isNumeric(Request.QueryString("TID")) Then lngTopicID = LngC(Request.QueryString("TID")) Else lngTopicID = 0

'If there is no Topic ID then redirect the user to the main forum page
If lngTopicID = 0 Then

	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("default.asp" & strQsSID1)
End If



'Read in the post ID if we are going straight to a post
If isNumeric(Request.QueryString("PID")) Then lngGetPostID = LngC(Request.QueryString("PID"))
	
'Fix for stupid IE6 redirect bug (IE SUCKS!!!) MS have fixed this for IE7
If lngGetPostID = 0 AND Request.QueryString("PID") <> "" Then
	'Remove the '#' hash anchor mark IE 6 includes in the string when doing a redirect (IE SUCKS!!)
	If InStr(Request.QueryString("PID"), "#") Then lngGetPostID = CLng(Mid(Request.QueryString("PID"), 1, InStr(Request.QueryString("PID"), "#")-1))
End If





'Sort the direction of db results from session
If blnMobileBrowser Then 'Do not reverse order for mobile view
	strSortBy = "ASC"
ElseIf UCase(Request.QueryString("OB") = "DESC") Then
	strSortBy = "DESC"
ElseIf UCase(Request.QueryString("OB") = "ASC") Then
	strSortBy = "ASC"
ElseIf getSessionItem("OB") = "DESC" Then
	strSortBy = "DESC"
ElseIf getSessionItem("OB") = "ASC" Then
	strSortBy = "ASC"
Else
	strSortBy = strDefaultPostOrder
End If

'Save the sort results order to app session
If Request.QueryString("OB") <> "" AND getSessionItem("OB") <> strSortBy Then Call saveSessionItem("OB", strSortBy)






'If this is a top priority post across all forums then read in the forum ID form the querystring and ingnore the real topic forum ID
If Request.QueryString("PR") = "3" Then intForumID = IntC(Request.QueryString("FID"))


strSQL = "" & _
"SELECT" & strDBTop1 & " " & strDbTable & "Category.Cat_ID, " & strDbTable & "Category.Cat_name, " & strDbTable & "Forum.Forum_ID AS ForumID, " & strDbTable & "Forum.Sub_ID, " & strDbTable & "Forum.Forum_name, " & strDbTable & "Forum2.Forum_name AS Main_forum, " & strDbTable & "Forum.Password, " & strDbTable & "Forum.Forum_code, " & strDbTable & "Forum.Locked AS forumLocked, " & strDbTable & "Topic.Poll_ID, " & strDbTable & "Topic.Subject, " & strDbTable & "Topic.Icon, " & strDbTable & "Topic.Locked AS topicLocked, " & strDbTable & "Topic.Priority, " & strDbTable & "Topic.No_of_views, " & strDbTable & "Topic.Hide, " & strDbTable & "Topic.Event_date, " & strDbTable & "Topic.Event_date_end, " & strDbTable & "Topic.Rating, " & strDbTable & "Topic.Rating_Votes, " & strDbTable & "Permissions.*  " & _
"FROM " & strDbTable & "Category" & strDBNoLock & ", " & strDbTable & "Forum" & strDBNoLock & ", " & strDbTable & "Forum AS " & strDbTable & "Forum2" & strDBNoLock & ", " & strDbTable & "Permissions" & strDBNoLock & ", " & strDbTable & "Topic" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Category.Cat_ID = " & strDbTable & "Forum.Cat_ID " & _
	"AND " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Permissions.Forum_ID "
'Get the forum ID from the topic table
If intForumID = 0 Then
	strSQL = strSQL  & _
	"AND " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Topic.Forum_ID "
'Else this is a priority post so get the forum ID from the querystring (security built in later to make sure the user is not trying to view a post they shouldn't)
Else
	strSQL = strSQL & "" & _
	"AND " & strDbTable & "Forum.Forum_ID = " & intForumID & " "

End If
	strSQL = strSQL & "" & _
	"AND (" & strDbTable & "Forum.Forum_ID = " & intForumID & " OR " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Topic.Forum_ID) " & _
	"AND (" & strDbTable & "Forum.Sub_ID = " & strDbTable & "Forum2.Forum_ID OR (" & strDbTable & "Forum.Sub_ID = 0 AND " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Forum2.Forum_ID)) " & _
	"AND " & strDbTable & "Topic.Topic_ID = " & lngTopicID & " " & _
	"AND (" & strDbTable & "Permissions.Author_ID = " & lngLoggedInUserID & " OR " & strDbTable & "Permissions.Group_ID = " & intGroupID & ") " & _
"ORDER BY " & strDbTable & "Forum.Sub_ID, " & strDbTable & "Forum.Forum_Order, " & strDbTable & "Permissions.Author_ID DESC" & strDBLimit1 & ";"


'Set error trapping
On Error Resume Next

'Query the database
rsCommon.Open strSQL, adoCon

'If an error has occurred write an error to the page
If Err.Number <> 0 Then	Call errorMsg("An error has occurred while executing SQL query on database.", "get_forum_data", "forum_posts.asp")

'Disable error trapping
On Error goto 0


'If there is no record returended then set a message to say that
If rsCommon.EOF Then

	'If there are no thread's to display then display the appropriate error message
	strSubject = strTxtNoThreads



'Else get the details of the forum, permissions, and topic details
Else

	'Read in forum details from the database
	intForumID = Cint(rsCommon("ForumID"))
	intCatID = CInt(rsCommon("Cat_ID"))
	strCatName = rsCommon("Cat_name")
	strForumName = rsCommon("Forum_name")
	strMasterForumName = rsCommon("Main_forum")
	intMasterForumID = CLng(rsCommon("Sub_ID"))
	blnForumLocked = CBool(rsCommon("forumLocked"))
	blnHideTopic = CBool(rsCommon("Hide"))
	If isDate(rsCommon("Event_date")) Then dtmEventDate = CDate(rsCommon("Event_date"))
	If isDate(rsCommon("Event_date_end")) Then dtmEventDateEnd = CDate(rsCommon("Event_date_end"))

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
	strForumPassword = rsCommon("Password")

	'Read in the topic details
	lngPollID = CLng(rsCommon("Poll_ID"))
	strSubject = rsCommon("Subject")
	strTopicIcon = rsCommon("Icon")
	blnTopicLocked = CBool(rsCommon("topicLocked"))
	intPriority = CInt(rsCommon("Priority"))
	If isNull(rsCommon("Rating")) Then dblTopicRating = 0 Else dblTopicRating = CDbl(rsCommon("Rating"))
	If isNull(rsCommon("Rating_Votes")) Then lngTopicVotes = 0 Else lngTopicVotes = CDbl(rsCommon("Rating_Votes"))
		
		
		
	'If the title is not included in the link redirect page back to itself with the title included
	If Request.QueryString("title") <> "" AND Request.QueryString("title") <> SeoUrlTitle(strSubject, "")  Then
		
		'Clean up
		rsCommon.Close
		Call closeDatabase()
		
		'Redirect uback to this file with the correct page na
		'If not on page 1 include page number
		If intRecordPositionPageNum > 1 Then
			Response.Redirect("forum_posts.asp?TID=" & lngTopicID & "&PN=" & intRecordPositionPageNum & SeoUrlTitle(strSubject, "&title="))
			Response.End
		Else
			Response.Redirect("forum_posts.asp?TID=" & lngTopicID & SeoUrlTitle(strSubject, "&title="))
			Response.End
		End If
		
			
	End If	
		
		

	'Get the keywords	
	If blnDynamicMetaTags Then strDynamicKeywords = dynamicKeywords(strSubject)

	'Clean up input to prevent XXS hack
	strSubject = formatInput(strSubject)
	
	'If the user has no read writes then kick them out
	If blnRead = False Then

		'Reset Server Objects
		rsCommon.Close
		Call closeDatabase()


		'Redirect to a page asking for the user to enter the forum password
		Response.Redirect("insufficient_permission.asp" & strQsSID1)
	End If


	'If the forum requires a password and a logged in forum code is not found on the users machine then send them to a login page
	If strForumPassword <> "" AND (getCookie("fID", "Forum" & intForumID) <> rsCommon("Forum_code") AND getSessionItem("FP" & intForumID) <> "1") Then

		'Reset Server Objects
		rsCommon.Close
		Call closeDatabase()

		'Redirect to a page asking for the user to enter the forum password
		Response.Redirect("forum_password_form.asp?RP=PT&FID=" & intForumID & "&TID=" & lngTopicID & strQsSID3)
	End If



	'Update no. of views, if not a bot (bots can index a page 100's of times thus making the view count unreliable)
	If NOT strOSType = "Search Robot" AND NOT strOSType = "Unknown" Then

		'Read in the number of views for the page form the database
		lngNumberOfViews = CLng(rsCommon("No_of_views"))

		'Add 1 to the number of views the Topic has had
		lngNumberOfViews = lngNumberOfViews + 1

		'Write the number of times the Topic has been viewed back to the database
		'Initalise the strSQL variable with the SQL string
		strSQL = "UPDATE " & strDbTable & "Topic" & strRowLock & " " & _
		"SET " & strDbTable & "Topic.No_of_views = " & lngNumberOfViews & " " & _
		"WHERE " & strDbTable & "Topic.Topic_ID = " & lngTopicID & ";"

		'Set error trapping
		On Error Resume Next

		'Write to the database
		adoCon.Execute(strSQL)

		'If an error has occurred write an error to the page
	  	If Err.Number <> 0 Then Call errorMsg("An error has occurred while writing to the database.", "update_no._views", "forum_posts.asp")

		'Disable error trapping
		On Error goto 0
	End If
End If

'clean up
rsCommon.Close


'Securtiy check - if this isn't a Priority post and the user is just trying to view a post in a forum they don't have permission to
'then reload the page without the forum ID
If Request.QueryString("PR") = "3" AND NOT intPriority = 3 Then

	'Reset Server Objects
	Call closeDatabase()

	'Redirect back to this page without the forum ID
	Response.Redirect("forum_posts.asp?TID=" & lngTopicID & strQsSID3)
End If



'Use the application session to pass around what forum this user is within
If IntC(getSessionItem("FID")) <> intForumID Then Call saveSessionItem("FID", intForumID)






'********************************
'****   Get Topics SQL Query ****
'********************************


'Start with WHERE Cluases as these are used in both the count query and in the main query

strSQLFromWhere = strSQLFromWhere & "" & _
"WHERE " & strDbTable & "Thread.Topic_ID = " & lngTopicID & " "
'If this isn't a moderator only display hidden posts if the user posted them
If blnModerator = false AND blnAdmin = false Then
	strSQLFromWhere = strSQLFromWhere & "AND (" & strDbTable & "Thread.Hide = " & strDBFalse & " "
	'Don't display hidden posts if guest
	If intGroupID <> 2 Then strSQLFromWhere = strSQLFromWhere & "OR " & strDbTable & "Thread.Author_ID = " & lngLoggedInUserID
	strSQLFromWhere = strSQLFromWhere & ") "
End If





'Count the number of posts in this topic
'***************************************	
strSQL = "" & _
"SELECT Count(" & strDbTable & "Thread.Thread_ID) AS PostCount " & _
"FROM " & strDbTable & "Thread" & strDBNoLock & " "
strSQL = strSQL & strSQLFromWhere & ";"


'Set error trapping
On Error Resume Next
		
'Query the database
rsCommon.Open strSQL, adoCon
	
'If an error has occurred write an error to the page
If Err.Number <> 0 Then	Call errorMsg("An error has occurred while executing SQL query on database.", "post_count", "forum_posts.asp")
				
'Disable error trapping
On Error goto 0
	
'Read in member count from database
lngTotalRecords = CLng(rsCommon("PostCount"))
	
'Close recordset
rsCommon.close



'Calculate what posts to get
'***************************

'Count the number of pages for the topics using '\' so that any fraction is omitted
lngTotalRecordsPages = lngTotalRecords \ intThreadsPerPage
	
'If there is a remainder or the result is 0 then add 1 to the total num of pages
If lngTotalRecords Mod intThreadsPerPage > 0 OR lngTotalRecordsPages = 0 Then lngTotalRecordsPages = lngTotalRecordsPages + 1
	
'Get the record poistion to display from
'If the page number to show from is higher than the last page number then the last page number is the highest page
If (intRecordPositionPageNum > lngTotalRecordsPages) OR (Request.QueryString("get") = "last") Then
	
	'Set the page position number to the highest page number
	intRecordPositionPageNum = lngTotalRecordsPages
End If




'Jump straight to a post (eg. Get Post 35 (&PID=35#35))
'************************************************
	
'If we need to jump straight to a particular post calculate which page to go to
If lngGetPostID <> 0 AND lngTotalRecordsPages > 1 Then
	
	
	'Count the number of posts before the post we want
	strSQL = "" & _
	"SELECT Count(" & strDbTable & "Thread.Thread_ID) AS PrePostCount " &_
	"FROM " & strDbTable & "Thread" & strDBNoLock & " "
	strSQL = strSQL & strSQLFromWhere & " "
	If strSortBy = "DESC" Then
		strSQL = strSQL &"AND " & strDbTable & "Thread.Thread_ID > " & lngGetPostID & ";"
	Else
		strSQL = strSQL &"AND " & strDbTable & "Thread.Thread_ID < " & lngGetPostID & ";"
	End If
	
	
	'Set error trapping
	On Error Resume Next
			
	'Query the database
	rsCommon.Open strSQL, adoCon
		
	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then	Call errorMsg("An error has occurred while executing SQL query on database.", "get_last_post_count", "forum_posts.asp")
					
	'Disable error trapping
	On Error goto 0
		
	'Read in member count from database
	lngPrePostCount = CLng(rsCommon("PrePostCount"))
		
	'Close recordset
	rsCommon.close
	
	'Calculate page number (use FIX to get the whole number)
	intRecordPositionPageNum = FIX(lngPrePostCount / intThreadsPerPage) + 1
	
End If





'Start position
intStartPosition = ((intRecordPositionPageNum - 1) * intThreadsPerPage)





'Query to get Posts
'******************

'Intilise SQL query to get all the posts
'Use a LEFT JOIN for the Guest name as there may not be a Guest name and so we want to include null values
strSQL = "" & _
"SELECT "

'If SQL server advanced paging
If strDatabaseType = "SQLServer" AND blnSqlSvrAdvPaging Then
	strSQL = strSQL & " * " & _
	"FROM (SELECT "
	'Using TOP really speeds thing up on the first pages, but once you get up to really large number of pages it can cause order problems
	If intThreadsPerPage * intRecordPositionPageNum < 500 Then strSQL = strSQL & " TOP " & intThreadsPerPage * intRecordPositionPageNum  & " "
End If

'SQL SELECT
strSQL = strSQL & strDbTable & "Thread.Thread_ID, " & strDbTable & "Thread.Message, " & strDbTable & "Thread.Message_date, " & strDbTable & "Thread.Show_signature, " & strDbTable & "Thread.IP_addr, " & strDbTable & "Thread.Hide, " & strDbTable & "Thread.Thanks, " & strDbTable & "Thread.Answer, " & strDbTable & "Author.Author_ID, " & strDbTable & "Author.Username, " & strDbTable & "Author.Homepage, " & strDbTable & "Author.Location, " & strDbTable & "Author.No_of_posts, " & strDbTable & "Author.Points, " & strDbTable & "Author.Join_date, " & strDbTable & "Author.Signature, " & strDbTable & "Author.Active, " & strDbTable & "Author.Avatar, " & strDbTable & "Author.Avatar_title, " & strDbTable & "Author.Banned, " & strDbTable & "Author.Author_email, " & strDbTable & "Group.Name AS GroupName, " & strDbTable & "Group.Stars, " & strDbTable & "Group.Custom_stars, " & strDbTable & "GuestName.Name AS GuestName, " & strDbTable & "Group.Signatures "

'If SQL Server advanced paging
If strDatabaseType = "SQLServer"  AND blnSqlSvrAdvPaging Then
	strSQL = strSQL & ", ROW_NUMBER() OVER (ORDER BY " & strDbTable & "Thread.Message_date " & strSortBy & ") AS RowNum "

End If

'FROM Clause
strSQL = strSQL & "FROM (" & strDbTable & "Group INNER JOIN (" & strDbTable & "Author INNER JOIN " & strDbTable & "Thread ON " & strDbTable & "Author.Author_ID = " & strDbTable & "Thread.Author_ID) ON " & strDbTable & "Group.Group_ID = " & strDbTable & "Author.Group_ID) LEFT JOIN " & strDbTable & "GuestName ON " & strDbTable & "Thread.Thread_ID = " & strDbTable & "GuestName.Thread_ID "

'Get the FROM and WHERE cluases from earlier
strSQL = strSQL & strSQLFromWhere

'If SQL Server advanced paging
If strDatabaseType = "SQLServer" AND blnSqlSvrAdvPaging Then
	strSQL = strSQL & ") AS PagingQuery WHERE RowNum BETWEEN " & intStartPosition + 1 & " AND " & intStartPosition + intThreadsPerPage & " "

'Else Order by clause here
Else
	strSQL = strSQL & "ORDER BY " & strDbTable & "Thread.Message_date " & strSortBy & " "
End If


'mySQL limit operator
If strDatabaseType = "mySQL" Then
	strSQL = strSQL & "LIMIT " & intStartPosition & ", " & intThreadsPerPage
End If

strSQL = strSQL & ";"


'Set error trapping
On Error Resume Next



'Query the database
rsCommon.Open strSQL, adoCon

'If an error has occurred write an error to the page
If Err.Number <> 0 Then	Call errorMsg("An error has occurred while executing SQL query on database.", "get_topic_posts_data", "forum_posts.asp")

'Disable error trapping
On Error goto 0

'If there is a topic in the database then get the post data
If NOT rsCommon.EOF Then
	

	'Read in the topic recordset into an array
	sarryPosts = rsCommon.GetRows()
	
	'Clean up
	rsCommon.Close

	'SQL Query Array Look Up table
	'0 = tblThread.Thread_ID,
	'1 = tblThread.Message,
	'2 = tblThread.Message_date,
	'3 = tblThread.Show_signature,
	'4 = tblThread.IP_addr,
	'5 = tblThread.Hide,
	'6 = tblThread.Thanks,
	'7 = tblThread.Answer,
	'8 = tblAuthor.Author_ID,
	'9 = tblAuthor.Username,
	'10 = tblAuthor.Homepage,
	'11 = tblAuthor.Location,
	'12 = tblAuthor.No_of_posts,
	'13 = tblAuthor.Points,
	'14 = tblAuthor.Join_date,
	'15 = tblAuthor.Signature,
	'16 = tblAuthor.Active,
	'17 = tblAuthor.Avatar,
	'18 = tblAuthor.Avatar_title,
	'19 = tblAuthor.Banned,
	'20 = tblAuthor.Author_email
	'21 = tblGroup.Name,
	'22 = tblGroup.Stars,
	'23 = tblGroup.Custom_stars
	'23 = tblGuestName.Name
	'25 = tblGroup.Signatures
	
	
	
	
	'If advanced paging then workout the end and start position differently
	If (strDatabaseType = "SQLServer" AND blnSqlSvrAdvPaging) OR strDatabaseType = "mySQL" Then
		
		'End Position
		intEndPosition = Ubound(sarryPosts,2) + 1
	
		'Get the start position
		intCurrentRecord = 0
	
	'Else for slower paging for Access and SQL Server 2000	
	Else
		'End Position
		intEndPosition = intStartPosition + intThreadsPerPage
	
		'Get the start position
		intCurrentRecord = intStartPosition
	End If
	
	'Get meta description
	strMetaDescription = Trim(removeHTML(sarryPosts(1,0), 100, true))

	'Get the author ID of the first post (this is used for checking if the topic is hidden)
	lngUserID = CLng(sarryPosts(8,0))


'Else if no record is retuned we do not want to be displaying a foum title
Else

	strSubject = ""
	rsCommon.Close
End If




'Hidden Topic check (only display hidden tiopics to the admin/mods, or if the poster created the topic)
If blnHideTopic AND blnModerator = false AND blnAdmin = false Then
	If lngUserID <> lngLoggedInUserID OR intGroupID = 2 Then
		strSubject = ""
		lngTotalRecords = 0
		lngTotalRecordsPages = 0
	
	End If
End If



'*** Anwser Posts ****
'Read in the anwser posts, if we are on page 1, there is a post, and anwsers posts in on
If intRecordPositionPageNum = 1 AND lngTotalRecords > 2 AND NOT strAnswerPosts = "Off" AND strSortBy = "ASC" Then
	
	strSQL = "" & _
	"SELECT  " & strDbTable & "Thread.Thread_ID, " & strDbTable & "Thread.Message, " & strDbTable & "Thread.Message_date, " & strDbTable & "Thread.Show_signature, " & strDbTable & "Thread.IP_addr, " & strDbTable & "Thread.Hide, " & strDbTable & "Thread.Thanks, " & strDbTable & "Thread.Answer, " & strDbTable & "Author.Author_ID, " & strDbTable & "Author.Username, " & strDbTable & "Author.Homepage, " & strDbTable & "Author.Location, " & strDbTable & "Author.No_of_posts, " & strDbTable & "Author.Points, " & strDbTable & "Author.Join_date, " & strDbTable & "Author.Signature, " & strDbTable & "Author.Active, " & strDbTable & "Author.Avatar, " & strDbTable & "Author.Avatar_title, " & strDbTable & "Author.Banned, " & strDbTable & "Author.Author_email, " & strDbTable & "Group.Name AS GroupName, " & strDbTable & "Group.Stars, " & strDbTable & "Group.Custom_stars, " & strDbTable & "GuestName.Name AS GuestName, " & strDbTable & "Group.Signatures "  & _
	"FROM (" & strDbTable & "Group INNER JOIN (" & strDbTable & "Author INNER JOIN " & strDbTable & "Thread ON " & strDbTable & "Author.Author_ID = " & strDbTable & "Thread.Author_ID) ON " & strDbTable & "Group.Group_ID = " & strDbTable & "Author.Group_ID) LEFT JOIN " & strDbTable & "GuestName ON " & strDbTable & "Thread.Thread_ID = " & strDbTable & "GuestName.Thread_ID "  & _
	strSQLFromWhere & _
	" AND " & strDbTable & "Thread.Answer = " & strDBTrue & " " & _
	"ORDER BY " & strDbTable & "Thread.Message_date " & strSortBy & ";"
	
	'Set error trapping
	On Error Resume Next
	
	'Query the database
	rsCommon.Open strSQL, adoCon
	
	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then	Call errorMsg("An error has occurred while executing SQL query on database.", "get_answer_topic_posts_data", "forum_posts.asp")
	
	'Disable error trapping
	On Error goto 0
	
	'If there is a topic in the database then get the post data
	If NOT rsCommon.EOF Then
	
		'Read in the topic recordset into an array
		saryAnswerPosts = rsCommon.GetRows()
		
		'Get the number of results
		intTopTopicAnswers = Ubound(saryAnswerPosts,2)
	End If
	
	'Clean up
	rsCommon.Close
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


'See if there are any unread posts in this topic
If isArray(sarryUnReadPosts) Then
	For intUnReadForumPostsLoop = 0 to UBound(sarryUnReadPosts,2)
		'If an unread post is found in this topic 
		If CLng(sarryUnReadPosts(1,intUnReadForumPostsLoop)) = lngTopicID AND sarryUnReadPosts(3,intUnReadForumPostsLoop) = "1" Then 
			'Increament unread post count
			intUnReadPostCount = intUnReadPostCount + 1
			'Get first unread post ID
			If lngFirstUnreadPostID = 0 Then lngFirstUnreadPostID = CLng(sarryUnReadPosts(0,intUnReadForumPostsLoop))
		End If
	Next	
End If




'Page to link to for mutiple page (with querystrings if required)
strLinkPage = "forum_posts.asp?TID=" & lngTopicID & "&amp;"
If Request.QueryString("KW") <> "" Then strLinkPage = strLinkPage & "KW=" & Server.URLEncode(Request.QueryString("KW")) & "&amp;"
If Request.QueryString("FID") <> "" Then strLinkPage = strLinkPage & "FID=" & intForumID & "&amp;"
If Request.QueryString("PR") <> "" Then strLinkPage = strLinkPage & "PR=" & IntC(Request.QueryString("PR")) & "&amp;"
strLinkPageTitle = SeoUrlTitle(strSubject, "&amp;title=")




'If active users is enabled update the active users application array
If blnActiveUsers Then
	'If there are keywords this is a search
	If Request.QueryString("KW") <> "" Then
		'Call active users function
		saryActiveUsers = activeUsers(strTxtViewingTopic & " " & strTxtSearchResults, strSubject, strLinkPage & "PN=" & intRecordPositionPageNum & SeoUrlTitle(strSubject, "&amp;title="), intForumID)
	'Else normal topic
	Else
		'Call active users function
		saryActiveUsers = activeUsers(strTxtViewingTopic, strSubject, strLinkPage & "PN=" & intRecordPositionPageNum & SeoUrlTitle(strSubject, "&amp;title="), intForumID)
	End If
End If



'If URL Rewriting is enabled create the canonical to the page for improved SEO
If NOT Request.ServerVariables("HTTP_X_ORIGINAL_URL") = "" OR (NOT Request.ServerVariables("HTTP_X_REWRITE_URL") = "" AND InStr(Request.ServerVariables("HTTP_X_REWRITE_URL"), ".html") > 1) Then
	
	If intRecordPositionPageNum = 1 Then
		strCanonicalURL = strForumPath & SeoUrlTitle(strSubject, "") & "_topic" & lngTopicID & ".html"
	Else
		strCanonicalURL = strForumPath & SeoUrlTitle(strSubject, "") & "_topic" & lngTopicID & "_page" & intRecordPositionPageNum & ".html"
	End If
	
	'Create URL for liking on social networks
	strSocialNetButtonsURL = strForumPath & SeoUrlTitle(strSubject, "") & "_topic" & lngTopicID & ".html"

'Else canonical without URL rewriting
Else
	If intRecordPositionPageNum = 1 Then
		strCanonicalURL = strForumPath & "forum_posts.asp?TID=" & lngTopicID & SeoUrlTitle(strSubject, "&title=")
	Else
		strCanonicalURL = strForumPath & "forum_posts.asp?TID=" & lngTopicID & "&PN=" & intRecordPositionPageNum & SeoUrlTitle(strSubject, "&title=")
	End If	
	
	'Create URL for liking on social networks
	strSocialNetButtonsURL = strForumPath & "forum_posts.asp?TID=" & lngTopicID & SeoUrlTitle(strSubject, "&title=")
End If



'Get the session key
strFormID = getSessionItem("KEY")


'If an admin or moderator turn of the edit time frame
If blnAdmin or blnModerator Then intEditPostTimeFrame = 0


'Set bread crumb trail
'Display the category name
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & "<a href=""default.asp?C=" & intCatID & strQsSID2 & SeoUrlTitle(strCatName, "&title=") & """>" & strCatName & "</a>" & strNavSpacer

'Display if there is a main forum to the sub forums name
If intMasterForumID <> 0 Then strBreadCrumbTrail = strBreadCrumbTrail & "<a href=""forum_topics.asp?FID=" & intMasterForumID & strQsSID2 & SeoUrlTitle(strMasterForumName, "&title=") & """>" & strMasterForumName & "</a>" & strNavSpacer

'Display forum name
If strForumName = "" Then strBreadCrumbTrail = strBreadCrumbTrail &  strTxtNoForums Else strBreadCrumbTrail = strBreadCrumbTrail & "<a href=""forum_topics.asp?FID=" & intForumID & strQsSID2 & SeoUrlTitle(strForumName, "&title=")  & """>" & strForumName & "</a>"




'Set the status bar tools

'Active Topics Links
strStatusBarTools = strStatusBarTools & "&nbsp;&nbsp;<img src=""" & strImagePath & "active_topics." & strForumImageType & """ alt=""" & strTxtActiveTopics & """ title=""" & strTxtActiveTopics & """ style=""vertical-align: text-bottom"" /> <a href=""active_topics.asp" & strQsSID1 & """>" & strTxtActiveTopics & "</a>"

'If RSS XML enabled then display an RSS button to link to XML file
If blnRSS AND (isNull(strForumPassword) OR strForumPassword = "") Then strStatusBarTools = strStatusBarTools & " <a href=""RSS_post_feed.asp?TID=" & lngTopicID & SeoUrlTitle(strSubject, "&title=") & """ target=""_blank""><img src=""" & strImagePath & "rss." & strForumImageType & """ alt=""" & strTxtRSS & " - " & strSubject & """ title=""" & strTxtRSS & ": " & strSubject & """ /></a>"



'If there is nothing to display send a 404 header
If lngTotalRecords = 0 Then Response.Status = "404 Not Found"



'Write the HTML head of the page
%><!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title><% = strSubject & " - " & strMainForumName %><% If lngTotalRecordsPages > 1 Then Response.Write(" - " & strTxtPage & " " & intRecordPositionPageNum) %></title>
<meta name="generator" content="Web Wiz Forums <% = strVersion %>" /><% 

'No index if no post is availble
If lngTotalRecords = 0 Then Response.Write(vbCrLf & "<meta name=""robots"" content=""noindex"" />")

'Dynamic meta tags
If blnDynamicMetaTags AND lngTotalRecords > 0 Then 
%>
<meta name="description" content="<% = strMetaDescription %>" />
<meta name="keywords" content="<% = strDynamicKeywords & ", " & strBoardMetaKeywords %>" /><%

ElseIf lngTotalRecords > 0 Then
	
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
vbCrLf & "//-->" & vbCrLf)
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******


'If RSS Feed is enabled then have an alt link to the XML file for supporting browsers
If blnRSS AND lngTotalRecords > 0 AND (isNull(strForumPassword) OR strForumPassword = "") Then Response.Write(vbCrLf & "<link rel=""alternate"" type=""application/rss+xml"" title=""RSS 2.0"" href=""RSS_post_feed.asp?TID=" & lngTopicID & SeoUrlTitle(strSubject, "&title=") & """ />")

%>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" /><%


'Facebook Open Graph Protocol Meta Tags
If blnFacebookLike Then

	'Facebook Open Graph Protocol Meta Tags
	If strFacebookPageID <> "" Then Response.Write(vbCrLf & "<meta property=""fb:page_id"" content=""" & strFacebookPageID & """ />")
	If strFacebookImage <> "" Then Response.Write(vbCrLf & "<meta property=""og:image"" content=""" & strFacebookImage & """/>")

%>
<meta property="og:site_name" content="<% = strMainForumName %>"/>
<meta property="og:title" content="<% = strSubject %>"/>
<meta property="og:url" content=" <% = strCanonicalURL %>"/>
<meta property="og:type" content="article"/><%

End If


'Display javascript for form checking
If blnActiveMember AND blnForumLocked = false AND blnReply AND (intRecordPositionPageNum = lngTotalRecordsPages) AND blnQuickReplyForm Then

%>
<script language="JavaScript">
function CheckForm() {
	var errorMsg = '';
	var formArea = document.getElementById('frmMessageForm');
	
<%

	'If Gecko Madis API (RTE) need to strip default input from the API
	If RTEenabled = "Gecko" Then Response.Write(vbCrLf & "	//For Gecko Madis API (RTE)" & vbCrLf & "	if (formArea.message.value.indexOf('<br>') > -1 && formArea.message.value.length==6) formArea.message.value = '';" & vbCrLf)

	'If this is a guest posting check that they have entered their name
	If blnPost And lngLoggedInUserID = 2 Then
%>
	//Check for a name
	if (formArea.Gname.value==''){
		errorMsg += '\n<% =  strTxtNoNameError %>';
	}<%
	
	End If

	'If CAPTCHA is displayed check it's been entered
	If blnCAPTCHAsecurityImages AND lngLoggedInUserID = 2 Then 
	
	%>
	
	//Check for a security code
        if (formArea.securityCode.value == ''){
                errorMsg += "\n<% = strTxtErrorSecurityCode %>";
        }<%
	End If

%>

	//Check for message
	if (formArea.message.value==''){
		errorMsg += '\n<% =  strTxtNoMessageError %>';
	}
	
	//Check session is not expired
        if (formArea.session.value != "true"){
                errorMsg += "\n<% = strTxtWarningYourSessionHasExpiredRefreshPageFormDataWillBeLost %>";
        }

	//If there is aproblem with the form then display an error
	if (errorMsg != ''){
		msg = '<% =  strTxtErrorDisplayLine %>\n\n';
		msg += '<% =  strTxtErrorDisplayLine1 %>\n';
		msg += '<% =  strTxtErrorDisplayLine2 %>\n';
		msg += '<% =  strTxtErrorDisplayLine %>\n\n';
		msg += '<% =  strTxtErrorDisplayLine3 %>\n';

		errorMsg += alert(msg + errorMsg + '\n\n');
		return false;
	}
	return true;
}
</script><%

End If

'Show Google +1 button
If blnGooglePlusOne Then
%>   	
<script type="text/javascript" src="https://apis.google.com/js/plusone.js"></script><%

End If



If strVigLinkKey = "" Then 
	Response.Write(vbCrLf & " <script type=""text/javascript"" src=""https://syndication.webwiz.net/exped/?SKU=WWF11""></script>" & vbCrLf )


ElseIf NOT strVigLinkKey = "Disabled" Then 
	
	Response.Write(vbCrLf & "<script type=""text/javascript"">" & _
		vbCrLf & "    var vglnk = { key: '" & strVigLinkKey & "' };" & _
		vbCrLf & "  " & _
		vbCrLf & "    (function(d, t) {" & _
		vbCrLf & "      var s = d.createElement(t); s.type = 'text/javascript'; s.async = true;" & _
		vbCrLf & "      s.src = '//cdn.viglink.com/api/vglnk.js';" & _
		vbCrLf & "      var r = d.getElementsByTagName(t)[0]; r.parentNode.insertBefore(s, r);" & _
		vbCrLf & "    }(document, 'script'));" & _
		vbCrLf & "</script>" & _
		vbCrLf & "")
End If


%>
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<a name="top"></a>
<table class="basicTable" cellspacing="0" cellpadding="5" align="center">
 <tr>
  <td><h1><%
  	
'If the forum is locked show a locked pad lock icon
If blnForumLocked Then 
	Response.Write ("<img src=""" & strImagePath & "forum_lock." & strForumImageType & """ alt=""" & strTxtForumLocked & """ title=""" & strTxtForumLocked & """ style=""vertical-align: text-bottom"" />")
'If the topic is locked then have a locked icon
ElseIf blnTopicLocked Then 
	Response.Write ("<img src=""" & strImagePath & "topic_lock." & strForumImageType & """ alt=""" & strTxtTopicLocked & """ title=""" & strTxtTopicLocked & """ style=""vertical-align: text-bottom"" />")
End If

'If a calendar event then display so
If isDate(dtmEventDate) Then
	Response.Write(strSubject & " - " & strTxtEventDate & ": " & stdDateFormat(dtmEventDate, False))
	
	'If over multiple dates
	If isDate(dtmEventDateEnd) Then Response.Write(" - " & stdDateFormat(dtmEventDateEnd, False))
Else  	
 	Response.Write(strSubject) 
End If




%></h1></td>
 </tr>
</table>
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td>
   <a href="new_reply_form.asp?TID=<% = lngTopicID %>&amp;PN=<% = lngTotalRecordsPages %>&amp;TR=<% =  lngTotalRecords & strQsSID2 %>" title="<% = strTxtPostReply %>" class="largeButton" rel="nofollow">&nbsp;<% = strTxtPostReply %> <img src="<% = strImagePath %>reply.<% = strForumImageType %>" border="0" alt="<% = strTxtPostReply %>" /></a>
  </td>
  <td align="right"><!-- #include file="includes/page_link_inc.asp" --></td>
 </tr>
</table><%

'If there is a poll then display the poll include
If lngPollID > 0 Then
	%><!--#include file="includes/poll_display_inc.asp" --><%
End If


'If there are posts display them
If lngTotalRecords > 0 Then

	'If mobile browser display different header
	If blnMobileBrowser Then

%>
<table class="tableBorder" align="center" cellspacing="1" cellpadding="3" style="table-layout: fixed;">
 <tr class="tableLedger">
  <td>
   <a href="forum_posts.asp?TID=<% = lngTopicID & strQsSID2 & SeoUrlTitle(strSubject, "&title=") %>"><% = strSubject %></a> 
  </td>
 </tr><%
	'Else standard header layout
	Else

%>
<table class="tableBorder" align="center" cellspacing="1" cellpadding="10" style="table-layout: fixed;">
 <tr class="tableLedger">
  <td style="width:140px;"><% = strTxtAuthor %></td>
  <td>
   <div style="float:left;"><a href="forum_posts.asp?TID=<% = lngTopicID %>&amp;OB=DESC<% = strQsSID2 %>" title="<% = strTxtReverseSortOrder %>"><% = strTxtMessage %></a><% If strSortBy = "DESC" Then Response.Write(" <a href=""forum_posts.asp?TID=" & lngTopicID & "&amp;OB=ASC" & strQsSID2 & """><img src=""" & strImagePath & strSortBy & "." & strForumImageType & """ title=""" & strTxtReverseSortOrder & """ alt=""" & strTxtReverseSortOrder & """ /></a>") %><%
   	
   	'Display link to first unread post
	If intUnReadPostCount > 0 Then
		Response.Write(" / <a href=""forum_posts.asp?TID=" & lngTopicID)
		If intPriority = 3 Then Response.Write("&amp;FID=" & intForumID & "&amp;PR=3") 
		Response.Write("&amp;PID=" & lngFirstUnreadPostID & "#" & lngFirstUnreadPostID & strQsSID2 & """  title=""" & intUnReadPostCount & " " & strTxtNewPosts & """>" & strTxtViewUnreadPost & "</a>") 
   	End If
   	
   	%></div>
   <div style="float:right;" class="downDropParent"><%
   
   	'Topic rating if enabled
   	If blnTopicRating Then
   		
   		'If not a geust and not banned display a drop down link
   		If blnGuest = False AND blnBanned = False AND blnActiveMember Then
   			Response.Write("&nbsp;&nbsp;<span id=""ajaxRating""><span id=""TopicRatingLink"" onclick=""showDropDown('TopicRatingLink', 'dropDownTopicRating', 190, 0);"" class=""dropDownPointer"">")
   		End If	
   			
	   	'Show the present rating
	   	If dblTopicRating >= 1 Then
	   		Response.Write("&nbsp;" & strTxtRating & ": <img src=""" & strImagePath & Mid(CStr(dblTopicRating + 0.5), 1, 1) & "_star_topic_rating." & strForumImageType & """ alt=""" & strTxtTopicRating & ": " & lngTopicVotes & " " & strTxtVotes & ", " &  strTxtAverage & " " & FormatNumber(dblTopicRating, 2) & """ title=""" & strTxtTopicRating & ": " & lngTopicVotes & " " & strTxtVotes & ", " &  strTxtAverage & " " & FormatNumber(dblTopicRating, 2) & """ style=""vertical-align: text-bottom;"" /></span>") 
	   	
	   	'Ask if user wants to rate topic if not a guest
	   	ElseIf blnGuest = False AND blnBanned = False AND blnActiveMember Then
	   		Response.Write("<img src=""" & strImagePath & "topic_rating." & strForumImageType & """ alt=""" & strTxtTopicRating & """ title=""" & strTxtTopicRating & """ style=""vertical-align: text-bottom;"" />")
	   		Response.Write("&nbsp;" & strTxtRateTopic & "</span>") 
	   	End IF
	   	
	   	'If not guest complete drop down link
   		If blnGuest = False AND blnBanned = False AND blnActiveMember Then
	   		
   			'Display hidden div for ajax rating drop down
   			Response.Write("<iframe id=""dropDownTopicRating"" src=""rate_topic.asp?TID=" & lngTopicID & strQsSID2 & """ class=""dropDownTopicRating"" frameborder=""0"" scrolling=""no""></iframe></span>")	
   		End If
	End If
   		
	
	
'	'Share Topic Drop Down
'	If blnShareTopicLinks Then
'		Response.Write("&nbsp;&nbsp;<span id=""shareTopic"" onclick=""getAjaxData('ajax_topic_share.asp?TID=" & lngTopicID & "&URL=" & Server.URLEncode(strSocialNetButtonsURL) & "&Title=" & Server.URLEncode(strSubject) & strQsSID2 & "', 'showShareTopicMenu');showDropDown('shareTopic', 'shareTopicMenu', 400, 160);"" class=""dropDownPointer""><img src=""" & strImagePath & "openshare-icon." & strForumImageType & """ alt=""" & strTxtShareTopic & """ title=""" & strTxtShareTopic & """ style=""vertical-align: text-bottom"" /> " & strTxtShareTopic & "</span>" & _
'		"<div id=""shareTopicMenu"" class=""dropDownTopicShare""><span id=""showShareTopicMenu""></span></div>")
'	End If
'	
	
	'Topic Search Link
  	Response.Write("&nbsp;&nbsp;<span id=""SearchTopicLink"" onclick=""showDropDown('SearchTopicLink', 'dropDownTopicSearch', 230, 135);"" class=""dropDownPointer""><img src=""" & strImagePath & "search." & strForumImageType & """ alt=""" & strTxtTopic & " " & strTxtSearch & """ title=""" & strTxtTopic & " " & strTxtSearch & """ style=""vertical-align: text-bottom;"" />&nbsp;" & strTxtTopic & " " & strTxtSearch & "</span>") 
	
	
	
	
	'Modertor Tools
	If blnAdmin OR blnModerator Then
		Response.Write("&nbsp;&nbsp;<span id=""modTools"" onclick=""showDropDown('modTools', 'modToolsMenu', 120, 0);"" class=""dropDownPointer""><img src=""" & strImagePath & "moderator_tools." & strForumImageType & """ alt=""" & strTxtModeratorTools & """ title=""" & strTxtModeratorTools & """ style=""vertical-align: text-bottom"" /> " & strTxtModeratorTools & "</span>" & _
		"<div id=""modToolsMenu"" class=""dropDownMenu"">" & _
		"<a href=""pre_approved_topics.asp" & strQsSID1 & """><div>" & strTxtHiddenTopics & "</div></a>" & _
		"<a href=""javascript:winOpener('pop_up_topic_admin.asp?TID=" & lngTopicID & strQsSID2 & "','admin',1,1,600,285)""><div>" & strTxtTopicAdmin & "</div></a>")
		
		'Lock or un-lock forum if admin
		If blnTopicLocked Then
			Response.Write("<a href=""lock_topic.asp?mode=UnLock&amp;TID=" & lngTopicID & "&amp;XID=" & strFormID & strQsSID2 & """><div>" & strTxtUnLockTopic & "</div></a>")
		Else
			Response.Write("<a href=""lock_topic.asp?mode=Lock&amp;TID=" & lngTopicID & "&amp;XID=" & strFormID & strQsSID2 & """><div>" & strTxtLockTopic & "</div></a>")
		End If
		
		'Hide or show topic
		If blnHideTopic = false Then
			Response.Write("<a href=""lock_topic.asp?mode=Hide&amp;TID=" & lngTopicID & "&amp;XID=" & strFormID & strQsSID2 & """><div>" & strTxtHideTopic & "</div></a>")
		Else
			Response.Write("<a href=""lock_topic.asp?mode=Show&amp;TID=" & lngTopicID & "&amp;XID=" & strFormID & strQsSID2 & """><div>" & strTxtShowTopic & "</div></a>")
		End If
		
		Response.Write("<a href=""delete_topic.asp?TID=" & lngTopicID & "&amp;PN=" & intRecordPositionPageNum & "&amp;XID=" & strFormID & strQsSID2 & """ onclick=""return confirm('" & strTxtDeleteTopicAlert & "')""><div>" & strTxtDeleteTopic & "</div></a>")
		Response.Write("</div>")
	End If
	
	
	'Topic Options drop down
	Response.Write("&nbsp;&nbsp;<span id=""topicOptions"" onclick=""")
	'If we need a subscription link then include a call to the ajax function
	If intGroupID <> 2 AND blnEmail AND blnLoggedInUserEmail AND blnActiveMember Then Response.Write("getAjaxData('ajax_email_notify.asp?TID=" & lngTopicID & "&amp;PN=" & intRecordPositionPageNum & strQsSID2 & "', 'ajaxEmailSub');")
		
	Response.Write("showDropDown('topicOptions', 'optionsMenu', 122, 26);"" class=""dropDownPointer""><img src=""" & strImagePath & "topic_options." & strForumImageType & """ alt=""" & strTxtTopicOptions & """ title=""" & strTxtTopicOptions & """ style=""vertical-align: text-bottom;"" /> " & strTxtTopicOptions & "</span>" & _
	"<div id=""optionsMenu"" class=""dropDownMenu"">" & _
	"<a href=""new_reply_form.asp?TID=" & lngTopicID & "&amp;PN=" &  lngTotalRecordsPages & "&amp;TR=" &  lngTotalRecords & strQsSID2 & """ rel=""nofollow""><div>" & strTxtPostReply & "</div></a>" & _
	"<a href=""new_topic_form.asp?FID=" & intForumID & strQsSID2 & """ rel=""nofollow""><div>" & strTxtCreateNewTopic & "</div></a>")
	If blnPollCreate Then Response.Write("<a href=""new_poll_form.asp?FID=" & intForumID & strQsSID2 & """ rel=""nofollow""><div>" & strTxtCreateNewPoll & "</div></a>")
	Response.Write("<a href=""printer_friendly_posts.asp?TID=" & lngTopicID & strQsSID2 & """ target=""_blank""><div title=""" & strTxtPrintVersion & """> " & strTxtPrintVersion & "</div></a>" & _
	"<a href=""https://translate.google.com/translate?u=" & Server.URLEncode(strCanonicalURL) & """ target=""_blank""><div title=""" & strTxtTranslate & """> " & strTxtTranslate & " " & strTxtTopic & "</div></a>")
	'Display option to subscribe or un-subscribe to topic
	If intGroupID <> 2 AND blnEmail AND blnLoggedInUserEmail AND blnActiveMember Then Response.Write("<span id=""ajaxEmailSub""></span>")
	  
  
  
  %></div>
  </td>
 </tr><%
	End If

 	'Calculate the post number
	intThreadNo = (intRecordPositionPageNum - 1) * intThreadsPerPage


	'Do....While Loop to loop through the recorset to display the topic posts
	Do While intCurrentRecord < intEndPosition

		'If there are no post records left to display then exit loop
		If intCurrentRecord >= lngTotalRecords Then Exit Do
					
		
		'SQL Query Array Look Up table
		'0 = tblThread.Thread_ID,
		'1 = tblThread.Message,
		'2 = tblThread.Message_date,
		'3 = tblThread.Show_signature,
		'4 = tblThread.IP_addr,
		'5 = tblThread.Hide,
		'6 = tblThread.Thanks,
		'7 = tblThread.Answer,
		'8 = tblAuthor.Author_ID,
		'9 = tblAuthor.Username,
		'10 = tblAuthor.Homepage,
		'11 = tblAuthor.Location,
		'12 = tblAuthor.No_of_posts,
		'13 = tblAuthor.Points,
		'14 = tblAuthor.Join_date,
		'15 = tblAuthor.Signature,
		'16 = tblAuthor.Active,
		'17 = tblAuthor.Avatar,
		'18 = tblAuthor.Avatar_title,
		'19 = tblAuthor.Banned,
		'20 = tblAuthor.Author_email
		'21 = tblGroup.Name,
		'22 = tblGroup.Stars,
		'23 = tblGroup.Custom_stars
		'24 = tblGuestName.Name
		'25 = tblGroup.Signatures
		
	
		'Set to falese for top of page topic answer
		blnTopTopicAnswer = False
		
		
		'Calculate the thread number
		intThreadNo = intThreadNo + 1
		
		
		'If we are at the first record after topic starter see if we have an anwser post to display
		If intCurrentRecord = 1 AND intTopTopicAnswers => 0 AND intTopTopicAnswers => intCurrentAnswerRecord Then
			
			'Set Top Topic anser to true
			blnTopTopicAnswer = True
			
			'Read in threads details for the topic from the database
			lngMessageID = CLng(saryAnswerPosts(0,intCurrentAnswerRecord))
			strMessage = saryAnswerPosts(1,intCurrentAnswerRecord)
			dtmPostDate = CDate(saryAnswerPosts(2,intCurrentAnswerRecord))
			strAuthorIP = saryAnswerPosts(4,intCurrentAnswerRecord)
			blnHidePost = CBool(saryAnswerPosts(5,intCurrentAnswerRecord))
			
			If isNull(saryAnswerPosts(6,intCurrentAnswerRecord)) Then lngPostThanks = 0 Else lngPostThanks = CInt(saryAnswerPosts(6,intCurrentAnswerRecord))
			blnTopicAnswer = CBool(saryAnswerPosts(7,intCurrentAnswerRecord))
			
			lngUserID = CLng(saryAnswerPosts(8,intCurrentAnswerRecord))
			strUsername = saryAnswerPosts(9,intCurrentAnswerRecord)
	
			strAuthorHomepage = saryAnswerPosts(10,intCurrentAnswerRecord)
			strAuthorLocation = saryAnswerPosts(11,intCurrentAnswerRecord)
			lngAuthorNumOfPosts = CLng(saryAnswerPosts(12,intCurrentAnswerRecord))
			If isNull(saryAnswerPosts(13,intCurrentAnswerRecord)) Then lngAuthorNumOfPoints = 0 Else lngAuthorNumOfPoints = CLng(saryAnswerPosts(13,intCurrentAnswerRecord))
			dtmAuthorRegistration = CDate(saryAnswerPosts(14,intCurrentAnswerRecord))
			strAuthorSignature = saryAnswerPosts(15,intCurrentAnswerRecord)
			strAuthorEmail = saryAnswerPosts(20,intCurrentAnswerRecord)
			
	
			strAuthorAvatar = saryAnswerPosts(17,intCurrentAnswerRecord)
			strMemberTitle = saryAnswerPosts(18,intCurrentAnswerRecord)
			blnSuspended = CBool(saryAnswerPosts(19,intCurrentAnswerRecord))
			strGroupName = saryAnswerPosts(21,intCurrentAnswerRecord)
			intRankStars = CInt(saryAnswerPosts(22,intCurrentAnswerRecord))
			strRankCustomStars = saryAnswerPosts(23,intCurrentAnswerRecord)
			blnShowSignature = CBool(saryAnswerPosts(25,intCurrentAnswerRecord))
	
			strGuestUsername = saryAnswerPosts(24,intCurrentAnswerRecord)
			
			'Move to next anwser record
			intCurrentAnswerRecord = intCurrentAnswerRecord + 1
		Else

			'Initliase the unread post variable
			blnUnreadPost = False
			
			'If the message being display has the same number as the previous one, do not display
			'(This is used for Anwser posts so that the same post is not displayed twice)
			If CLng(sarryPosts(0,intCurrentRecord)) = lngMessageID Then intCurrentRecord = intCurrentRecord + 1
	
			'Read in threads details for the topic from the database
			lngMessageID = CLng(sarryPosts(0,intCurrentRecord))
			strMessage = sarryPosts(1,intCurrentRecord)
			dtmPostDate = CDate(sarryPosts(2,intCurrentRecord))
			If isNull(sarryPosts(4,intCurrentRecord)) Then strAuthorIP = "Unknown" Else strAuthorIP = sarryPosts(4,intCurrentRecord)
			blnHidePost = CBool(sarryPosts(5,intCurrentRecord))
			
			If isNull(sarryPosts(6,intCurrentRecord)) Then lngPostThanks = 0 Else lngPostThanks = CInt(sarryPosts(6,intCurrentRecord))
			blnTopicAnswer = CBool(sarryPosts(7,intCurrentRecord))
			
			lngUserID = CLng(sarryPosts(8,intCurrentRecord))
			strUsername = sarryPosts(9,intCurrentRecord)
	
			strAuthorHomepage = sarryPosts(10,intCurrentRecord)
			strAuthorLocation = sarryPosts(11,intCurrentRecord)
			lngAuthorNumOfPosts = CLng(sarryPosts(12,intCurrentRecord))
			If isNull(sarryPosts(13,intCurrentRecord)) Then lngAuthorNumOfPoints = 0 Else lngAuthorNumOfPoints = CLng(sarryPosts(13,intCurrentRecord))
			dtmAuthorRegistration = CDate(sarryPosts(14,intCurrentRecord))
			strAuthorSignature = sarryPosts(15,intCurrentRecord)
			strAuthorEmail = sarryPosts(20,intCurrentRecord)

	
			strAuthorAvatar = sarryPosts(17,intCurrentRecord)
			strMemberTitle = sarryPosts(18,intCurrentRecord)
			blnSuspended = CBool(sarryPosts(19,intCurrentRecord))
			strGroupName = sarryPosts(21,intCurrentRecord)
			intRankStars = CInt(sarryPosts(22,intCurrentRecord))
			strRankCustomStars = sarryPosts(23,intCurrentRecord)
			blnShowSignature = CBool(sarryPosts(25,intCurrentRecord))
	
			strGuestUsername = sarryPosts(24,intCurrentRecord)
		End If
		
		
		'If the users is suspended or signatures not allowed for group then blank their signature
		If blnSuspended OR blnShowSignature = False Then strAuthorSignature = ""
		
		
		'If there is a newer post than the last time the unread posts array was initilised run it again
		If dtmPostDate > CDate(Session("dtmUnReadPostCheck")) Then Call UnreadPosts()
		
		'Mark if this post has been read, if the unread posts array exists
		If isArray(sarryUnReadPosts) AND  dtmPostDate > dtmLastVisitDate Then
			
			'Loop through the unread posts array
			For intUnReadForumPostsLoop = 0 to UBound(sarryUnReadPosts,2)
				
				'Mark the post as read
				If CLng(sarryUnReadPosts(0,intUnReadForumPostsLoop)) = lngMessageID Then 
					
					If sarryUnReadPosts(3,intUnReadForumPostsLoop) = "1" Then blnUnreadPost = True
					
					'mark post as read
					sarryUnReadPosts(3,intUnReadForumPostsLoop) = "0"
					
					'save back to session
					
					'Save the array at applicion level
					If isArray(Application("sarryUnReadPosts" & strSessionID)) Then  
						Application.Lock
						Application("sarryUnReadPosts" & strSessionID) = sarryUnReadPosts 
						Application.UnLock
					'save the array at session level
					ElseIf isArray(Session("sarryUnReadPosts")) Then 
						Session("sarryUnReadPosts") = sarryUnReadPosts
					End If
					
					'Exit for to save some processing time
					Exit For
				End If
			Next	
		End If

		

		'Calculate the row colour
		If intThreadNo MOD 2 = 0 Then
			strPostTableRowClass = "msgEvenTableRow"
			strPostTableSideClass = "msgEvenTableSide"
			strPostTableTopClass = "msgEvenTableTop"
		Else
			strPostTableRowClass = "msgOddTableRow"
			strPostTableSideClass = "msgOddTableSide"
			strPostTableTopClass = "msgOddTableTop"
		End If
		
		'If an answer post
		If blnTopicAnswer Then 
			strPostTableTopClass = "msgAnswerTableTop"
		End If

		'If this is a hidden post then change the row colour to highlight it
		If blnHidePost Then strPostTableRowClass = "msgHiddenTableRow"
			


		'If the poster is a guest see if they have entered their name in the GuestName table and get it
		If lngUserID = 2 AND strGuestUsername <> "" Then strUsername = strGuestUsername


		'If the post contains a quote or code block then format it
		If InStr(1, strMessage, "[QUOTE=", 1) > 0 AND InStr(1, strMessage, "[/QUOTE]", 1) > 0 Then strMessage = formatUserQuote(strMessage)
		If InStr(1, strMessage, "[QUOTE]", 1) > 0 AND InStr(1, strMessage, "[/QUOTE]", 1) > 0 Then strMessage = formatQuote(strMessage)
		If InStr(1, strMessage, "[CODE]", 1) > 0 AND InStr(1, strMessage, "[/CODE]", 1) > 0 Then strMessage = formatCode(strMessage)
		If InStr(1, strMessage, "[HIDE]", 1) > 0 AND InStr(1, strMessage, "[/HIDE]", 1) > 0 Then strMessage = formatHide(strMessage)
		If InStr(1, strMessage, "[SPOILER]", 1) > 0 AND InStr(1, strMessage, "[/SPOILER]", 1) > 0 Then strMessage = formatSpoiler(strMessage, False)


		'If the Post or signature contains Flash or YouTube BBcodes then display them
		If blnFlashFiles Then
			'Flash
			If InStr(1, strMessage, "[FLASH", 1) > 0 AND InStr(1, strMessage, "[/FLASH]", 1) > 0 Then strMessage = formatFlash(strMessage)
			If InStr(1, strAuthorSignature, "[FLASH", 1) > 0 AND InStr(1, strAuthorSignature, "[/FLASH]", 1) > 0 Then strAuthorSignature = formatFlash(strAuthorSignature)
		End If
		
		'If YouTube
		If blnYouTube Then
			'YouTube
			If InStr(1, strMessage, "[TUBE]", 1) > 0 AND InStr(1, strMessage, "[/TUBE]", 1) > 0 Then strMessage = formatYouTube(strMessage)
			If InStr(1, strAuthorSignature, "[TUBE]", 1) > 0 AND InStr(1, strAuthorSignature, "[/TUBE]", 1) > 0 Then strAuthorSignature = formatYouTube(strAuthorSignature)
		End If


		'If the message has been edited parse the 'edited by' XML into HTML for the post
		If InStr(1, strMessage, "<edited>", 1) Then strMessage = editedXMLParser(strMessage)


		'Call the function to highlight search words if coming froma search page
		If strSearchKeywords <> "" Then strMessage = searchHighlighter(strMessage, sarySearchWord)
		
		
		'Calulate rowspan amount for table below side bar	
		If CBool(sarryPosts(3,intCurrentRecord)) AND strAuthorSignature <> "" AND blnSignatures Then 
			intPostRowsSpan = 3
		Else 
			intPostRowsSpan = 2
		End If
		

		'If not mobile browser
		If blnMobileBrowser Then
%>
<td valign="top" class="<% = strPostTableTopClass %>">
 <a name="<% = lngMessageID %>"></a><%
 	
 			'If an answer post
			If blnTopicAnswer Then
				Response.Write("<img src=""" & strImagePath & "answer.png"" title=""" & strAnswerPostsWording & """ alt=""" & strAnswerPostsWording & """ /> <strong>" & strAnswerPostsWording & "</strong><br />")
			End If
 			
 			'If unread post display an un-read post icon next to the post
			If blnUnreadPost Then 
				Response.Write(vbCrLf & "   <img src=""" & strImagePath & "new_post." & strForumImageType & """ alt=""" & strTxtNewPost & ": " & strTxtDirectLinkToThisPost & """ /> ")
			
			'Else display a bullet incon with direct link to post
			Else
				'Display direct link to post
				Response.Write(vbCrLf & "   <img src=""" & strImagePath & "bullet." & strForumImageType & """ alt=""" & strTxtDirectLinkToThisPost & """ /> ")
	     		End If
	     		
%>			
 <a href="member_profile.asp?PF=<% = lngUserID & strQsSID2 %>" rel="nofollow"><% = strUsername %></a> <%
   		
   		'Display message post date and time
		Response.Write(DateFormat(dtmPostDate) & " " & strTxtAt & " " & TimeFormat(dtmPostDate)) 
			
		'Display IP if admin or moderator
		If blnAdmin OR (blnModerator AND blnModViewIpAddresses) Then Response.Write(" - " & strTxtIP & ": " & strAuthorIP)
			
		'Disply post options if there are any for this user
	 	If (blnTopicLocked = False) OR (blnAdmin OR blnModerator) Then
	  		
		  	Response.Write(vbCrLf & "   <span style=""float:right"">")
		
			'If the logged in user is the person who posted the message or the forum administrator/moderator then allow them to edit or delete the message
			If (lngLoggedInUserID = lngUserID AND blnForumLocked = False AND blnActiveMember AND blnTopicLocked = False) OR (blnAdmin OR blnModerator) Then
			
					
				'Only let the user edit the post if they have edit rights
				If blnEdit OR blnAdmin Then
					Response.Write(vbCrLf & "    &nbsp;<img src=""" & strImagePath & "edit.png"" title=""" & strTxtEditPost & """ alt=""" & strTxtEditPost & """ /> <a href=""edit_post_form.asp?PID=" & lngMessageID & "&amp;PN=" & intRecordPositionPageNum)
					If lngPollID > 0 AND intThreadNo = 1 Then Response.Write("&amp;POLL=" & lngPollID)
					Response.Write("" & strQsSID2 & """ rel=""nofollow"">" & strTxtEdit & "</a>")
				End If
				
				'If you want delete button for mobile users add it here
			
			End If
			
			'If the topic is not locked put in a link for someone to quote this message
			 If blnTopicLocked = False AND blnPollNoReply = False AND blnHidePost = False Then 
			      			
				Response.Write(vbCrLf & "    &nbsp;<img src=""" & strImagePath & "quote.png"" title=""" & strTxtQuote & "&nbsp;" & strUsername & """ alt=""" & strTxtQuote & "&nbsp;" & strUsername & """ /> <a href=""new_reply_form.asp?Quote=1&amp;PID=" & lngMessageID & "&amp;PN=" & lngTotalRecordsPages & "&amp;TR=" & lngTotalRecords & strQsSID2 & """ rel=""nofollow"">" & strTxtQuote & "</a>" & _
				" &nbsp;<img src=""" & strImagePath & "reply.png"" title=""" & strTxtPostReply & """ alt=""" & strTxtPostReply & """ /> <a href=""new_reply_form.asp?PID=" & lngMessageID & "&amp;PN=" & lngTotalRecordsPages & "&amp;TR=" &  lngTotalRecords & strQsSID2 & """ rel=""nofollow"">" & strTxtReply & "</a>")
			      	 			
			End If
				
			Response.Write(vbCrLf & "   </span>")
		End If	
			
		
		
%>	
</td>
 </tr>
 <tr class="<% = strPostTableRowClass %>">
  <td valign="top" class="msgLineDevider">
   <!-- Start Member Post <% = strUsername %> -->
   <div class="msgBody">
   <% = strMessage %>
   </div>
   <!-- End Member Post <% = strUsername %> -->
  </td>
 </tr>
 <tr class="postSeparatorTableRow">
  <td align="right"><a href="#top"><img src="<% = strImagePath %>back_to_top.<% = strForumImageType %>" alt="<% = strTxtBackToTop %>" title="<% = strTxtBackToTop %>" /></a></td>
 </tr>
<%			
				
	
	
			
		'Else non mobile brosser view
		Else		
%>
 <tr>
  <td rowspan="<% = intPostRowsSpan %>" valign="top" class="<% = strPostTableSideClass %>">
   <a name="<% = lngMessageID %>"></a>
   <span id="userProfile<% = intThreadNo %>" onclick="showDropDown('userProfile<% = intThreadNo %>', 'profileMenu<% = intThreadNo %>', 160, 0);" class="msgSideProfile" title="<% = strTxtViewDropDown %>"><% = strUsername %> <img src="<% = strImagePath %>drop_down.<% = strForumImageType %>" alt="<% = strTxtViewDropDown %>" /></span>
   <div id="profileMenu<% = intThreadNo %>" class="dropDownMenu">
    <a href="member_profile.asp?PF=<% = lngUserID & strQsSID2 %>" rel="nofollow"><div><% = strTxtMembersProfile %></div></a><%

	    		'Memeber drop down options
	    		If blnPrivateMessages Then Response.Write(vbCrLf & "    <a href=""pm_new_message_form.asp?name=" & Server.URLEncode(Replace(strUsername, "'", "\'",  1, -1, 1)) & strQsSID2 & """ rel=""nofollow""><div>" & strTxtSendPrivateMessage & "</div></a>")
	    		Response.Write(vbCrLf & "    <a href=""search_form.asp?USR=" & Server.URLEncode(Replace(strUsername, "'", "\'",  1, -1, 1)) & strQsSID2 & """ rel=""nofollow""><div>" & strTxtFindMembersPosts & "</div></a>")
	    		If blnPrivateMessages Then Response.Write(vbCrLf & "    <a href=""pm_buddy_list.asp?name=" & Server.URLEncode(Replace(strUsername, "'", "\'",  1, -1, 1)) & strQsSID2 & """ rel=""nofollow""><div>" & strTxtAddToBuddyList & "</div></a>")
%> 
   </div>
   <br /><% = strGroupName %>
   <br /><img src="<% If strRankCustomStars <> "" Then Response.Write(strRankCustomStars) Else Response.Write(strImagePath & intRankStars & "_star_rating.png") %>" alt="<% = strGroupName %>" title="<% = strGroupName %>" />
   <br /><%


			
		        'If the user has an avatar then display it
		        If blnAvatar = True AND strAuthorAvatar <> "" Then
		        	
		        	'If SSL is enabled, but the avatar is not HTTPS then remove avatar
				If strSslEnabled = "Enabled" AND InStr(strAuthorAvatar, "http://") Then 
					strAuthorAvatar = ""
				'Display avatar
				Else
		        		Response.Write("<img src=""" & strAuthorAvatar & """ alt=""" & strTxtAvatar & """ class=""avatar"" onerror=""this.src='avatars/blank_avatar.jpg';"" />")
				End If
			End If
			
		       	'If there is a title for this member then display it
		       	If strMemberTitle <> "" Then Response.Write(vbCrLf & "   <br />" & strMemberTitle)
	
		        'If not a guest post then display some details
		        If lngUserID <> 2 Then
	
		         	'Show the joined date
		         	Response.Write(vbCrLf & "   <br /><br />" & strTxtJoined & ": " & DateFormat(dtmAuthorRegistration))
	
		         	'If the is a location display it
		         	If strAuthorLocation <> "" Then Response.Write(vbCrLf & "   <br />" & strTxtLocation & ": " & strAuthorLocation)
	
	
				'If active users is enabled see if the user is online
				If blnActiveUsers Then
	
					'Display if the user is online
					blnIsUserOnline = False
	
					'Loop through the active users array
					For intLoop = 1 To UBound(saryActiveUsers, 2)
						If saryActiveUsers(1, intLoop) = lngUserID Then blnIsUserOnline = True
					Next
	
					'Display if the user is online
					If blnIsUserOnline Then Response.Write(vbCrLf & "   <br />" & strTxtStatus & ": " & strTxtOnLine2) Else Response.Write(vbCrLf & "      <br />" & strTxtStatus & ": " & strTxtOffLine)
				End If
	
	
		         	'Display the num of posts
		         	Response.Write(vbCrLf & "   <br />" & strTxtPoints & ": " & lngAuthorNumOfPoints)
		         	
			End If
%>
  </td>
  <td valign="top" class="<% = strPostTableTopClass %>" style="height:15px;"><%
	
			'If an answer post
			If blnTopicAnswer Then
				Response.Write("<img src=""" & strImagePath & "answer.png"" title=""" & strAnswerPostsWording & """ alt=""" & strAnswerPostsWording & """ /> <span class=""lgText"">" & strAnswerPostsWording & "</span><br />")
			End If
			
	  		'Disply post options if there are any for this user
	 		If (blnTopicLocked = False) OR (blnAdmin OR blnModerator) Then
	  		
		  		Response.Write(vbCrLf & "   <span style=""float:right"">")
		  		
		  		'If post is hidden display approved post option
		  		If blnHidePost AND (blnAdmin OR blnModerator) Then Response.Write(vbCrLf & "   <img src=""" & strImagePath & "approve_post." & strForumImageType & """ title=""" & strTxtApprovePost & """ alt=""" & strTxtApprovePost & """ style=""vertical-align: text-bottom;"" /> <a href=""show_post.asp?PID=" & lngMessageID & "&amp;XID=" & strFormID & strQsSID2 & """>" & strTxtApprovePost & "</a>&nbsp;&nbsp;")
		  		
		  		
		  		'Post options drop down	
		  		Response.Write(vbCrLf & "    <span id=""postOptions" & intThreadNo & """ onclick=""showDropDown('postOptions" & intThreadNo & "', 'postMenu" & intThreadNo & "', 130, 0);"" class=""postOptions""><img src=""" & strImagePath & "post_options." & strForumImageType & """ alt=""" & strTxtPostOptions & """ title=""" & strTxtPostOptions & """ style=""vertical-align: text-bottom;"" /> " & strTxtPostOptions & "</span>" & _
		  		vbCrLf & "    <div id=""postMenu" & intThreadNo & """ class=""dropDownMenu"">")
				
				'If the topic is not locked put in a link for someone to quote this message
		      		If blnTopicLocked = False AND blnPollNoReply = False AND blnHidePost = False Then 
					Response.Write(vbCrLf & "     <a href=""new_reply_form.asp?TID=" & lngTopicID & "&amp;PN=" &  lngTotalRecordsPages & "&amp;TR=" &  lngTotalRecords & strQsSID2 & """ rel=""nofollow""><div>" & strTxtPostReply & "</div></a>" & _
					vbCrLf & "     <a href=""new_reply_form.asp?Quote=1&amp;PID=" & lngMessageID & "&amp;PN=" & lngTotalRecordsPages & "&amp;TR=" & lngTotalRecords & strQsSID2 & """ rel=""nofollow""><div>" & strTxtQuote & "&nbsp;" & strUsername & "</div></a>")
				End If
				
				'Display the report post feature if email is enabled.
		        	If blnEmail AND lngUserID <> lngLoggedInUserID AND blnActiveMember Then Response.Write(vbCrLf & "     <a href=""report_post.asp?PID=" & lngMessageID & "&amp;FID=" & intForumID & "&amp;TID=" & lngTopicID & "&amp;PN=" & intRecordPositionPageNum & strQsSID2 & """ rel=""nofollow""><div>" & strTxtReportPost & "</div></a> ")
				
				'If the logged in user is the person who posted the message or the forum administrator/moderator then allow them to edit or delete the message
			      	If (lngLoggedInUserID = lngUserID AND blnForumLocked = False AND blnActiveMember AND blnTopicLocked = False) OR (blnAdmin OR blnModerator) Then
		
			      	   	'Only let the user edit the post if they have edit rights
			      	   	If (blnEdit AND (intEditPostTimeFrame = 0 OR DateDiff("n", dtmPostDate, Now()) <= intEditPostTimeFrame)) OR (blnAdmin OR blnModerator) Then
			      	   		Response.Write(vbCrLf & "     <a href=""edit_post_form.asp?PID=" & lngMessageID & "&amp;PN=" & intRecordPositionPageNum)
			      	   		If lngPollID > 0 AND intThreadNo = 1 Then Response.Write("&amp;POLL=" & lngPollID)
			      	   		Response.Write("" & strQsSID2 & """><div>" & strTxtEditPost & "</div></a>")
					End If
		
					'Only let a normal member delete there post if someone hasn't posted a reply
		      			If (lngTotalRecords = intThreadNo) OR (blnAdmin OR blnModerator) Then
		
			            		'Only let the user delete the post if they have delete rights
			            		If blnDelete OR blnAdmin Then
			            			Response.Write(vbCrLf & "     <a href=""delete_post.asp?PID=" & lngMessageID & "&amp;PN=" & intRecordPositionPageNum & "&amp;XID=" & strFormID & strQsSID2 & """ OnClick=""return confirm('" & strTxtDeletePostAlert & "')""><div>" & strTxtDeletePost & "</div></a>")
			      			End If
			      		End If
			      	End If
		
		
		   	      	'If the user is the forum admin or a moderator
			      	If blnAdmin OR blnModerator Then
		
			      		'Let them hide the post
			      		If blnHidePost Then
			      			Response.Write(vbCrLf & "     <a href=""show_post.asp?PID=" & lngMessageID & "&amp;XID=" & strFormID & strQsSID2 & """><div>" & strTxtApprovePost & "</div></a>")
			      		Else
			      			Response.Write(vbCrLf & "     <a href=""hide_post.asp?PID=" & lngMessageID & "&amp;XID=" & strFormID & strQsSID2 & """ OnClick=""return confirm('" & strTxtAreYouSureYouWantToHidePost & "')""><div>" & strTxtHidePost & "</div></a>")
			      		End If
			      	
				End If
			      		
			      	'Set post as Anwser
				If ((strAnswerPosts = "admin" AND blnAdmin) OR (strAnswerPosts = "admin_mods" AND (blnAdmin OR blnModerator)) OR (strAnswerPosts = "all_members" AND blnGuest = False AND blnPost)) AND intCurrentRecord > 0 Then
					If blnTopicAnswer Then
						Response.Write("<a href=""answer_post.asp?mode=Remove&amp;PID=" & lngMessageID & "&amp;XID=" & strFormID & strQsSID2 & """><div>" & strTxtUnSetAsAnswer & "</div></a>")
					Else
						Response.Write("<a href=""answer_post.asp?mode=Set&amp;PID=" & lngMessageID & "&amp;XID=" & strFormID & strQsSID2 & """><div>" & strTxtSetAsAnswer & "</div></a>")
					End If
				End If
					
				'If the user is the forum admin or a moderator
			      	If blnAdmin OR blnModerator Then
		
			      		'Let them move the post to another forum
			      		Response.Write(vbCrLf & "     <a href=""javascript:winOpener('move_post_form.asp?PID=" & lngMessageID & strQsSID2 & "','move',1,1,580,265)""><div>" & strTxtMovePost & "</div></a>")
				
					'Display the IP address for IP blocking to admin/moderators
					If (blnAdmin OR (blnModerator AND blnModViewIpAddresses)) AND (strAuthorIP <> "") Then Response.Write(vbCrLf & "     <a href=""javascript:winOpener('pop_up_IP_blocking.asp?IP=" & strAuthorIP & strQsSID2 & "','ip',1,1,500,475)"" title=""" & strAuthorIP & """><div>" & strTxtBlockUsersIP & "</div></a>")
				
					'If StopForumSpam enabled then add a link for admins and mods if they aree allowed to see IP
					'If (blnAdmin OR (blnModerator AND blnModViewIpAddresses)) AND blnStopForumSpam AND strStopForumSpamApiKey <> "" Then Response.Write(vbCrLf & "     <a href=""http://www.stopforumspam.com/add?username=" & strUsername & "&email=" & strAuthorEmail & "&ip_addr=" & strAuthorIP & "&api_key=" & strStopForumSpamApiKey & """ target=""_blank""><div>" & strSubmitAsSpammer & "</div></a>")
				End If
				
				Response.Write(vbCrLf & "    </div>")
				


		
		                'If the topic is not locked put in a link for someone to quote this message
		      		If blnTopicLocked = False AND blnPollNoReply = False AND blnHidePost = False Then 
		      			
		      			'If thanks is enabled
		      			If blnPostThanks Then
						'Thanks drown down to the right
				  		Response.Write(vbCrLf & "    &nbsp;&nbsp;<span id=""postThanks" & intThreadNo & """ onclick=""getAjaxData('ajax_post_thanks.asp?PID=" & lngMessageID & "&amp;XID=" & strFormID & strQsSID2 & "', 'ajaxThanks" & intThreadNo & "');showDropDown('postThanks" & intThreadNo & "', 'postThanksMsg" & intThreadNo & "', 180, 0);"" class=""postThanks""><img src=""" & strImagePath & "thanks." & strForumImageType & """ title=""" & strTxtThanks & " (" & lngPostThanks & ")"" alt=""" & strTxtThanks & " (" & lngPostThanks & ")"" style=""vertical-align: text-bottom;"" /> " & strTxtThanks & "(" & lngPostThanks & ")</span>" & _
				  		vbCrLf & "    <div id=""postThanksMsg" & intThreadNo & """ class=""dropDownMenu"">" & _
						vbCrLf & "     <span id=""ajaxThanks" & intThreadNo & """  class=""postThanksDropDown""></span>" & _
						vbCrLf & "    </div>")
			      		End If
		      			
		      			Response.Write(vbCrLf & "    &nbsp;&nbsp;<img src=""" & strImagePath & "quote." & strForumImageType & """ title=""" & strTxtQuote & "&nbsp;" & strUsername & """ alt=""" & strTxtQuote & "&nbsp;" & strUsername & """ style=""vertical-align: text-bottom;"" /> <a href=""new_reply_form.asp?Quote=1&amp;PID=" & lngMessageID & "&amp;PN=" & lngTotalRecordsPages & "&amp;TR=" & lngTotalRecords & strQsSID2 & """ rel=""nofollow"">" & strTxtQuote & "</a>&nbsp;&nbsp;" & _
		      			"<img src=""" & strImagePath & "reply." & strForumImageType & """ title=""" & strTxtPostReply & """ alt=""" & strTxtPostReply & """ style=""vertical-align: text-bottom;"" /><a href=""new_reply_form.asp?PID=" & lngMessageID & "&amp;PN=" & lngTotalRecordsPages & "&amp;TR=" &  lngTotalRecords & strQsSID2 & """ rel=""nofollow"">" & strTxtReply & "</a>")
		      		
		      		
		      		'ElseIf thanks is enabled and not includeed in the above
				ElseIf blnPostThanks Then
					
					'Thanks drown down to the left
			  		Response.Write(vbCrLf & "    &nbsp;&nbsp;<span id=""postThanks" & intThreadNo & """ onclick=""getAjaxData('ajax_post_thanks.asp?PID=" & lngMessageID & "&amp;XID=" & strFormID & strQsSID2 & "', 'ajaxThanks" & intThreadNo & "');showDropDown('postThanks" & intThreadNo & "', 'postThanksMsg" & intThreadNo & "', 180, 106);"" class=""postThanks""><img src=""" & strImagePath & "thanks." & strForumImageType & """ title=""" & strTxtThanks & " (" & lngPostThanks & ")"" alt=""" & strTxtThanks & " (" & lngPostThanks & ")"" style=""vertical-align: text-bottom;"" /> " & strTxtThanks & "(" & lngPostThanks & ")</span>" & _
			  		vbCrLf & "    <div id=""postThanksMsg" & intThreadNo & """ class=""dropDownMenu"">" & _
					vbCrLf & "     <span id=""ajaxThanks" & intThreadNo & """  class=""postThanksDropDown""></span>" & _
					vbCrLf & "    </div>")
		      		End If
		      		
		      		Response.Write(vbCrLf & "   </span>")
	
			End If
	
			
			'If unread post display an un-read post icon next to the post
			If blnUnreadPost Then 
				Response.Write(vbCrLf & "   <a href=""forum_posts.asp?TID=" & lngTopicID & "&amp;PID=" & lngMessageID & strQsSID2 & SeoUrlTitle(strSubject, "&amp;title=") & "#" & lngMessageID & """ title=""" & strTxtNewPost & ": " & strTxtDirectLinkToThisPost & """><img src=""" & strImagePath & "new_post." & strForumImageType & """ alt=""" & strTxtNewPost & ": " & strTxtDirectLinkToThisPost & """ /></a> ")
			
			'Else display a bullet incon with direct link to post
			Else
				'Display direct link to post
				Response.Write(vbCrLf & "   <a href=""forum_posts.asp?TID=" & lngTopicID & "&amp;PID=" & lngMessageID & strQsSID2 & SeoUrlTitle(strSubject, "&amp;title=") & "#" & lngMessageID & """ title=""" & strTxtDirectLinkToThisPost & """><img src=""" & strImagePath & "bullet." & strForumImageType & """ alt=""" & strTxtDirectLinkToThisPost & """ /></a> ")
	     		End If
	     		
	     		'If first record display topic title
	     		If intRecordPositionPageNum = 1 AND intCurrentRecord = 0 Then 
	     			Response.Write("<strong>")
				'If a calendar event then display so
				If isDate(dtmEventDate) Then Response.Write(strTxtCalendarEvent & ": ") Else Response.Write(strTxtTopic & ": ") 
				Response.Write(strSubject & "</strong><br />&nbsp;&nbsp;&nbsp;&nbsp;")
			End If
	
			'Display message post date and time
			Response.Write(strTxtPosted & " " & DateFormat(dtmPostDate) & " " & strTxtAt & " " & TimeFormat(dtmPostDate)) 
			
			'Display IP if admin or moderator
			If blnAdmin OR (blnModerator AND blnModViewIpAddresses) Then Response.Write(" - " & strTxtIP & ": " & strAuthorIP & " <a href=""https://network-tools.webwiz.net/ip-information.htm?ip=" & Server.URLEncode(strAuthorIP) & """ target=""_blank""><img src=""" & strImagePath & "new_window.png"" alt=""" & strTxtIP & " " & strTxtInformation & """ title=""" & strTxtIP & " " & strTxtInformation & """ /></a>")
%></td>
 </tr>
 <tr class="<% = strPostTableRowClass %>" style="height:200px;min-height:200px;">
  <td valign="top" class="msgLineDevider">
   <!-- Start Member Post <% = strUsername %> -->
   <div class="msgBody">
   <% = strMessage %>
   </div>
   <!-- End Member Post <% = strUsername %> --><%

			'If the user wants there signature the display it
			If CBool(sarryPosts(3,intCurrentRecord)) AND strAuthorSignature <> "" AND blnSignatures Then
				Response.Write(vbCrLf & "  </td>" & _
				vbCrLf & " </tr>" & _
				vbCrLf & " <tr class=""" & strPostTableRowClass & """>" & _
				vbCrLf & "  <td class=""msgLineDevider"">" & _
				formatSignature(strAuthorSignature))
			End If
%>
  </td>
 </tr><%
 
			'Inject ad in to page if enabled
			If intCurrentRecord = 0 AND (strForumPostAd <> "") Then
				
				strPostTableRowClass = "msgEvenTableRow"
				strPostTableSideClass = "msgEvenTableSide"
				strPostTableTopClass = "msgEvenTableTop"
		
%>
 <tr class="postSeparatorTableRow">
  <td colspan="2" align="right"><a href="#top"><img src="<% = strImagePath %>back_to_top.<% = strForumImageType %>" alt="<% = strTxtBackToTop %>" title="<% = strTxtBackToTop %>" /></a></td>
 </tr>
 <tr>
  <td valign="top" class="<% = strPostTableSideClass %>">
   <span class="msgSideProfile"><% = strTxtSponsor %></span>
  </td>
  <td class="<% = strPostTableRowClass %>" align="center" nowrap="nowrap"><br /><%			
		
				If strForumPostAd <> "" Then Response.Write(strForumPostAd)
%><br /><br />
</td>
 </tr><%
			
			End If
			
		
 %>
 <tr class="postSeparatorTableRow">
  <td colspan="2" align="right"><a href="#top"><img src="<% = strImagePath %>back_to_top.<% = strForumImageType %>" alt="<% = strTxtBackToTop %>" title="<% = strTxtBackToTop %>" /></a></td>
 </tr><%
		End If

 		'Move to the next record
		If blnTopTopicAnswer = False Then intCurrentRecord = intCurrentRecord + 1
	Loop
%>
</table>
<iframe id="dropDownTopicSearch" src="quick_topic_search.asp?TID=<% = lngTopicID & strQsSID2 %>" class="dropDownTopicSearch" frameborder="0" scrolling="no" rel="nofollow"></iframe><%


'Else display there are no posts to display so display error message
Else

%>
<br />
<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
 <tr>
  <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtError %></strong></td>
 </tr>
 <tr>
  <td><% = strTxtNoThreads %></td>
 </tr>
<table><%



End If



%>
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td>
   <a href="new_reply_form.asp?TID=<% = lngTopicID %>&amp;PN=<% = lngTotalRecordsPages %>&amp;TR=<% =  lngTotalRecords & strQsSID2 %>" title="<% = strTxtPostReply %>" class="largeButton" rel="nofollow">&nbsp;<% = strTxtPostReply %> <img src="<% = strImagePath %>reply.<% = strForumImageType %>" border="0" alt="<% = strTxtPostReply %>" /></a>
  </td>
  <td align="right" nowrap="nowrap"><%

'Display table for email and/or page links
If lngTotalRecordsPages > 1 Then  
  
  	%><!-- #include file="includes/page_link_inc.asp" --><%

End If
%>
  </td>
 </tr>
 <tr>
  <td><%

'Share Topic Drop Down
If blnShareTopicLinks AND blnMobileBrowser = False Then
	Response.Write("&nbsp;&nbsp;<span id=""shareTopic"" onclick=""getAjaxData('ajax_topic_share.asp?TID=" & lngTopicID & "&URL=" & Server.URLEncode(strSocialNetButtonsURL) & "&Title=" & Server.URLEncode(strSubject) & strQsSID2 & "', 'showShareTopicMenu');showDropDown('shareTopic', 'shareTopicMenu', 720, 0);"" class=""dropDownPointer""><img src=""" & strImagePath & "share_button.png"" alt=""" & strTxtShareTopic & """ title=""" & strTxtShareTopic & """ style=""vertical-align: text-bottom"" onmouseover=""fadeImage(this)"" onmouseout=""unFadeImage(this)"" vspace=""4"" /></span>" & _
	"<div id=""shareTopicMenu"" class=""dropDownTopicShare""><span id=""showShareTopicMenu""></span></div> &nbsp;&nbsp; ")
End If


'Show twitter tweet button
If blnTwitterTweet Then
%>
  <script src="https://platform.twitter.com/widgets.js" type="text/javascript"></script>
  <a href="https://twitter.com/share?url=<% = Server.URLEncode(strSocialNetButtonsURL) %>&text=<% = Server.HTMLEncode(strSubject) %>" class="twitter-share-button">Tweet</a><%

End If


'Show Google +1 button
If blnGooglePlusOne Then

%>   	
   <g:plusone size="medium" href="<% = Server.URLEncode(strSocialNetButtonsURL) %>"></g:plusone><%

End If


'Show facebook like button
If blnFacebookLike Then
%>
   <iframe src="https://www.facebook.com/plugins/like.php?app_id=230119966998116&amp;href=<% = Server.URLEncode(strSocialNetButtonsURL) %>&amp;send=true&amp;layout=button_count&amp;width=400&amp;show_faces=false&amp;action=like&amp;colorscheme=light&amp;font=arial&amp;height=20" scrolling="no" frameborder="0" style="border:none; overflow:hidden; width:130px; height:20px;" allowTransparency="true"></iframe><%

End If




%></td>
 </tr>	
</table><%







'Quick Reply !!!!
If blnQuickReplyForm AND blnBannedIP = false AND blnActiveMember AND blnBanned = False AND blnForumLocked = false AND blnReply AND blnTopicLocked = false AND blnPollNoReply = false AND intRecordPositionPageNum = lngTotalRecordsPages AND blnMobileBrowser = False Then

	
	'Check to see if the user has email notification for this topic
	If blnEmail AND blnLoggedInUserEmail Then
		strSQL = "SELECT " & strDbTable & "EmailNotify.Author_ID  " & _
		"FROM " & strDbTable & "EmailNotify" & strRowLock & " " & _
		"WHERE " & strDbTable & "EmailNotify.Author_ID = " & lngLoggedInUserID & " AND " & strDbTable & "EmailNotify.Topic_ID = " & lngTopicID & ";"
		
		'Query the database
		rsCommon.Open strSQL, adoCon
		
		'If a record is returned then users has email notification enabled
		If NOT rsCommon.EOF Then blnReplyNotify = True
		
		'Close RS	
		rsCommon.Close
	End If
	
	
	
	%><!--#include file="includes/quick_reply_form_inc.asp" --><%
End If




%>
<br />
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td valign="top">
   <!-- #include file="includes/forum_jump_inc.asp" -->
  </td>
  <td align="right" class="smText" nowrap="nowrap"><!-- #include file="includes/show_forum_permissions_inc.asp" --></td>
 </tr>
</table>
<br />
<div align="center"><%

'Reset Server Objects
Call closeDatabase()

'If a mobile browser display an option to switch to and from mobile view
If blnMobileBrowser Then 
	Response.Write (strTxtViewIn & ": <strong>" & strTxtMoble & "</strong> | <a href=""" & strLinkPage & "MobileView=off" & strQsSID2 & """ rel=""nofollow"">" & strTxtClassic & "</a><br /><br />")
ElseIf blnMobileClassicView Then
	Response.Write (strTxtViewIn & ": <a href=""" & strLinkPage & "MobileView=on" & strQsSID2 & """ rel=""nofollow"">" & strTxtMoble & "</a> | <strong>" & strTxtClassic & "</strong><br /><br />")
End If


'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnLCode Then
	If blnTextLinks Then
		Response.Write("<span class=""text"" style=""font-size:10px"">Forum Software by <a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank""  style=""font-size:10px"">Web Wiz Forums&reg;</a> version " & strVersion & "</span>")
	Else
  		Response.Write("<a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank"" ><img src=""webwizforums_image.asp"" border=""0"" title=""Forum Software by Web Wiz Forums&reg; version " & strVersion& """ alt=""Forum Software by Web Wiz Forums&reg; version " & strVersion& """/></a>")
	End If

	Response.Write("<br /><span class=""text"" style=""font-size:10px"">Copyright &copy;2001-2019 Web Wiz Ltd.</span>")
End If
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******



'Display the process time
If blnShowProcessTime Then Response.Write("<span class=""smText""><br /><br />" & strTxtThisPageWasGeneratedIn & " " & FormatNumber(Timer() - dblStartTime, 3) & " " & strTxtSeconds & "</span>")


Response.Write("</div>")



'Display an alert message if the user is watching this topic for email notification
If Request.QueryString("EN") = "TS" Then
	Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
		"alert('" & strTxtYouWillNowBeNotifiedOfAllReplies & "');" & _
		"</script>")
End If

'Display an alert message if the user is not watching this topic for email notification
If Request.QueryString("EN") = "TU" Then
	Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
		"alert('" & strTxtYouWillNowNOTBeNotifiedOfAllReplies & "');" & _
		"</script>")
End If

'Display an alert message if the user is posting in a moderated forum
If Request.QueryString("MF") = "Y" Then
	Response.Write(vbCrLf & "<script language=""JavaScript"">" & _
		"alert('" & strTxtYouArePostingModeratedForum & " \n\n " & strTxtBeforePostDisplayedAuthorised & "');" & _
		"</script>")
End If

%>
<!-- #include file="includes/footer.asp" -->