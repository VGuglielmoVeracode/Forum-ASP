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




'Set the response buffer to true as we maybe redirecting and setting a cookie
Response.Buffer = true

'Make sure this page is not cached
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"



'Dimension variables
Dim sarryForums			'Holds the recordset array for all the categories and forums
Dim saryMemebrStats		'Holds the member stats
Dim strCategory			'Holds the category name
Dim intCatID			'Holds the id for the category
Dim strForumName		'Holds the forum name
Dim strForumDiscription		'Holds the forum description
Dim strForumPassword		'Holds the forum password if there is one
Dim lngNumberOfTopics		'Holds the number of topics in a forum
Dim lngNumberOfPosts		'Holds the number of Posts in the forum
Dim lngTotalNumberOfTopics	'Holds the total number of topics in a forum
Dim lngTotalNumberOfPosts	'Holds the total number of Posts in the forum
Dim intNumberofForums		'Holds the number of forums
Dim lngLastEntryMeassgeID	'Holds the message ID of the last entry
Dim dtmLastEntryDate		'Holds the date of the last entry to the forum
Dim strLastEntryUser		'Holds the the username of the user who made the last entry
Dim lngLastEntryUserID		'Holds the ID number of the last user to make and entry
Dim dtmLastEntryDateAllForums	'Holds the date of the last entry to all fourms
Dim strLastEntryUserAllForums	'Holds the the username of the user who made the last entry to all forums
Dim lngLastEntryUserIDAllForums	'Holds the ID number of the last user to make and entry to all forums
Dim blnForumLocked		'Set to true if the forum is locked
Dim intForumColourNumber	'Holds the number to calculate the table row colour
Dim blnHideForum		'Set to true if this is a hidden forum
Dim intCatShow			'Holds the ID number of the category to show if only showing one category
Dim intActiveUsers		'Holds the number of active users
Dim intActiveGuests		'Holds the number of active guests
Dim intActiveSearchRobots
Dim intActiveMembers		'Holds the number of logged in active members
Dim strMembersOnline		'Holds the names of the members online
Dim intSubForumID		'Holds the sub forum ID number
Dim strSubForumName		'Holds the sub forum name
Dim strSubForums		'Holds if there are sub forums
Dim dtmLastSubEntryDate		'Holds the date of the last entry to the forum
Dim strLastSubEntryUser		'Holds the the username of the user who made the last entry
Dim lngLastSubEntryUserID	'Holds the ID number of the last user to make and entry
Dim lngSubForumNumberOfPosts	'Holds the number of posts in the subforum
Dim lngSubForumNumberOfTopics	'Holds the number of topics in the subforum
Dim strSubForumPassword		'Holds sub forum password
Dim lngTotalRecords		'Holds the number of records
Dim intCurrentRecord		'Holds the current record position
Dim intTempRecord		'Holds a temporary record position for looping through records for any checks
Dim blnSubRead			'Holds if the user has entry to the sub forum
Dim lngNoOfMembers		'Holds the number of forum members
Dim intArrayPass		'Active users array counter
Dim strBirthdays		'String containing all those with birtdays today
Dim dtmNow			'Now date with off-set
Dim intBirtdayLoopCounter	'Holds the bitrhday loop counter
Dim intLastForumEntryID		'Holds the last forum ID for the last entry for link in forum stats
Dim intTotalViewingForum	'Holds the number of people viewing the forum, including sub forums
Dim intAnonymousMembers		'Holds the number of intAnonymous members online
Dim intUnReadPostCount		'Holds the count for the number of unread posts in the forum
Dim intUnReadForumPostsLoop	'Loop to count the number of unread posts in a forum
Dim lngTopicID			'Holds the topic ID
Dim strSubject			'Holds the subject
Dim lngSubTopicID		'Holds the topic ID
Dim strSubSubject		'Holds the subject
Dim strNewPostText
Dim strPageQueryString		'Holds the querystring for the page
Dim strForumImageIcon		'Hold an image icon for the forum
Dim strForumURL			'Holds the forum URL if a forum link
Dim dblGetLastPostID



'Initialise variables
lngTotalNumberOfTopics = 0
lngTotalNumberOfPosts = 0
intNumberofForums = 0
intForumColourNumber = 0
intActiveMembers = 0
intActiveGuests = 0
intActiveUsers = 0
intAnonymousMembers = 0
lngTotalRecords = 0
lngNoOfMembers = 0
intBirtdayLoopCounter = 0
intActiveSearchRobots = 0
dblGetLastPostID = CDbl(Now())




'Read in the qerystring
strPageQueryString = Request.QueryString()


'Remove the page title from the querystring beofre doing the sql injection test
If Request.QueryString("title") <> "" Then strPageQueryString = Replace(strPageQueryString, Request.QueryString("title"), "")

'Test querystrings for any SQL Injection keywords
Call SqlInjectionTest(strPageQueryString)



'Read in the category to show
If IsNumeric(Request.QueryString("C")) Then
	intCatShow = IntC(Request.QueryString("C"))
Else
	intCatShow = 0
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





'Read the various categories, forums, and permissions from the database in one hit for extra performance
'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "" & _
"SELECT " & strDbTable & "Category.Cat_ID, " & strDbTable & "Category.Cat_name, " & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Sub_ID, " & strDbTable & "Forum.Forum_name, " & strDbTable & "Forum.Forum_description, " & strDbTable & "Forum.No_of_topics, " & strDbTable & "Forum.No_of_posts, " & strDbTable & "Author.Username, " & strDbTable & "Forum.Last_post_author_ID, " & strDbTable & "Forum.Last_post_date, " & strDbTable & "Forum.Password, " & strDbTable & "Forum.Locked, " & strDbTable & "Forum.Hide, " & strDbTable & "Permissions.View_Forum, " & strDbTable & "Forum.Last_topic_ID, " & strDbTable & "Topic.Subject, " & strDbTable & "Forum.Forum_icon, " & strDbTable & "Forum.Forum_URL " & _
"FROM (((" & strDbTable & "Category INNER JOIN " & strDbTable & "Forum ON " & strDbTable & "Category.Cat_ID = " & strDbTable & "Forum.Cat_ID) LEFT JOIN " & strDbTable & "Topic ON " & strDbTable & "Forum.Last_topic_ID = " & strDbTable & "Topic.Topic_ID) INNER JOIN " & strDbTable & "Author ON " & strDbTable & "Forum.Last_post_author_ID = " & strDbTable & "Author.Author_ID) INNER JOIN " & strDbTable & "Permissions ON " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Permissions.Forum_ID " & _
"WHERE (" & strDbTable & "Permissions.Author_ID = " & lngLoggedInUserID & " OR " & strDbTable & "Permissions.Group_ID = " & intGroupID & ") " & _
"ORDER BY " & strDbTable & "Category.Cat_order, " & strDbTable & "Forum.Forum_Order, " & strDbTable & "Permissions.Author_ID DESC;"


'Set error trapping
On Error Resume Next
	
'Query the database
rsCommon.Open strSQL, adoCon

'If an error has occurred write an error to the page
If Err.Number <> 0 Then	Call errorMsg("An error has occurred while executing SQL query on database.", "get_forum_data", "default.asp")
			
'Disable error trapping
On Error goto 0


'Place the recordset into an array
If NOT rsCommon.EOF Then 
	sarryForums = rsCommon.GetRows()
	lngTotalRecords = Ubound(sarryForums,2) + 1
	
	'If showing a category check thnat the cat name has not been changed in the page name
	If intCatShow > 0 Then
		
		'Loop through categories to get teh cat name
		Do While intCurrentRecord <= Ubound(sarryForums,2)
			
			'If this is the category get the cat name
			If CInt(sarryForums(0,intCurrentRecord)) = intCatShow Then strCategory = sarryForums(1,intCurrentRecord)
			
			'Move to the next database record
			intCurrentRecord = intCurrentRecord + 1
		Loop
	
		'Reset loop counter
		intCurrentRecord = 0
		
		
		'If the title is not included in the link redirect page back to itself with the title included
		If Request.QueryString("title") <> "" AND Request.QueryString("title") <> SeoUrlTitle(strCategory, "")  Then
			
			'Clean up
			rsCommon.Close
			Call closeDatabase()
		
			'Redirect uback to this file with the correct page name
			Response.redirect("default.asp?C=" & intCatShow & SeoUrlTitle(strCategory, "&title="))
			Response.End
			
		End If	
			
	End If
End If

'Close the recordset
rsCommon.Close


'SQL Query Array Look Up table
'0 = Cat_ID
'1 = Cat_name
'2 = Forum_ID
'3 = Sub_ID
'4 = Forum_name
'5 = Forum_description
'6 = No_of_topics
'7 = No_of_posts
'8 = Last_post_author
'9 = Last_post_author_ID
'10 = Last_post_date
'11 = Password
'12 = Locked
'13 = Hide
'14 = Read 
'15 = Last_topic_ID
'16 = Topic.Subject
'17 = Forum_icon
'18 = Forum_URL




'Get the last signed up user and member stats and birthdays for use at bottom of page
If blnDisplayTodaysBirthdays Then
	
	'Get the now date with time off-set
	dtmNow = getNowDate()
	
	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "" & _
	"SELECT "
	If strDatabaseType = "SQLServer" OR strDatabaseType = "Access" Then
		strSQL = strSQL & " TOP 50 "
	End If
	strSQL = strSQL & _
	strDbTable & "Author.Username, " & strDbTable & "Author.Author_ID, " & strDbTable & "Author.DOB " & _
	"FROM " & strDbTable & "Author" & strDBNoLock & " " & _
	"WHERE MONTH(" & strDbTable & "Author.DOB) = " & Month(dtmNow) & " " & _
		"AND DAY(" & strDbTable & "Author.DOB) = " & Day(dtmNow) & " " & _
		"AND Active = " & strDBTrue & " " & _
		"AND Banned = " & strDBFalse & " " & _
	"ORDER BY " & strDbTable & "Author.Author_ID DESC "
	'mySQL limit operator
	If strDatabaseType = "mySQL" Then
		strSQL = strSQL & " LIMIT 50"
	End If
	strSQL = strSQL & ";"
	
	'Set error trapping
	On Error Resume Next
		
	'Query the database
	rsCommon.Open strSQL, adoCon
	
	'If an error has occurred write an error to the page
	If Err.Number <> 0 AND  strDatabaseType = "mySQL" Then	
		Call errorMsg("An error has occurred while executing SQL query on database.<br />Please check that the MySQL Server version is 4.1 or above.", "get_birthdays", "default.asp")
	ElseIf Err.Number <> 0 Then	
		Call errorMsg("An error has occurred while executing SQL query on database.", "get_birthdays", "default.asp")
	End If
	
				
	'Disable error trapping
	On Error goto 0
	
	'Place the recordset into an array
	If NOT rsCommon.EOF Then 
	
		'Read the recordset into an array
		saryMemebrStats = rsCommon.GetRows()
			
		'Loop through to get all members with birthdays today
		Do While intBirtdayLoopCounter <= Ubound(saryMemebrStats, 2)
			
			'If bitrhday is found for this date then add it to string
			If Month(dtmNow) = Month(saryMemebrStats(2, intBirtdayLoopCounter)) AND Day(dtmNow) = Day(saryMemebrStats(2, intBirtdayLoopCounter)) Then 
					
				'If there is already one birthday then place a comma before the next
				If strBirthdays <> "" Then strBirthdays = strBirthdays & ", "
					
				'Place the birthday into the Birthday array
				strBirthdays = strBirthdays & "<a href=""member_profile.asp?PF=" & saryMemebrStats(1, intBirtdayLoopCounter) & strQsSID2 &  """>" & saryMemebrStats(0, intBirtdayLoopCounter) & "</a> (" & Fix(DateDiff("m", saryMemebrStats(2, intBirtdayLoopCounter), Year(dtmNow) & "-" & Month(dtmNow) & "-" & Day(dtmNow))/12) & ")"
			End If
			
			'Increment loop counter by 1
			intBirtdayLoopCounter = intBirtdayLoopCounter + 1
		Loop
	End If
	
	'Close recordset
	rsCommon.close
End If






'Read in some stats about the last members
strSQL = "SELECT " & strDBTop1 & " " & strDbTable & "Author.Username, " & strDbTable & "Author.Author_ID " 
If NOT strDatabaseType = "mySQL" Then strSQL = strSQL & ", (SELECT COUNT (*) FROM "  & strDbTable & "Author WHERE 1 = 1) AS memberCount "
strSQL = strSQL & _
"FROM " & strDbTable & "Author" & strDBNoLock & " " & _
"WHERE Active = " & strDBTrue & " " & _
	"AND Banned = " & strDBFalse & " " & _
"ORDER BY " & strDbTable & "Author.Author_ID DESC " & strDBLimit1 & ";"

'Set error trapping
On Error Resume Next
	
'Query the database
rsCommon.Open strSQL, adoCon

'If an error has occurred write an error to the page
If Err.Number <> 0 Then	Call errorMsg("An error has occurred while executing SQL query on database.", "get_last_USR_+_count", "default.asp")
			
'Disable error trapping
On Error goto 0

'Place the recordset into an array
If NOT rsCommon.EOF Then 
	
	'Read in member count from database (if NOT mySQL)
	If NOT strDatabaseType = "mySQL" Then lngNoOfMembers = CLng(rsCommon("memberCount"))
	
	'Read the recordset into an array
	saryMemebrStats = rsCommon.GetRows()
End If

'Close recordset
rsCommon.close




'We have tgo use a seporate query to count the number of members in mySQL
If strDatabaseType = "mySQL" Then

	'Count the number of members
	strSQL = "SELECT Count(" & strDbTable & "Author.Author_ID) AS memberCount " & _
	"FROM " & strDbTable & "Author;"
	
	'Set error trapping
	On Error Resume Next
		
	'Query the database
	rsCommon.Open strSQL, adoCon
	
	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then	Call errorMsg("An error has occurred while executing SQL query on database.", "member_count", "default.asp")
				
	'Disable error trapping
	On Error goto 0
	
	'Read in member count from database
	lngNoOfMembers = CLng(rsCommon("memberCount"))
	
	'Close recordset
	rsCommon.close
End If






'If active users is enabled update the active users application array
If blnActiveUsers Then
	'Call active users function
	saryActiveUsers = activeUsers("", strTxtForumIndex, "default.asp", 0)
End If




'Set the status bar tools
'Active Topics Links
strStatusBarTools = strStatusBarTools & "&nbsp;&nbsp;<img src=""" & strImagePath & "active_topics." & strForumImageType & """ alt=""" & strTxtActiveTopics & """ title=""" & strTxtActiveTopics & """ style=""vertical-align: text-bottom"" /> <a href=""active_topics.asp" & strQsSID1 & """>" & strTxtActiveTopics & "</a> "
strStatusBarTools = strStatusBarTools & "&nbsp;&nbsp;<img src=""" & strImagePath & "unanswered_topics." & strForumImageType & """ alt=""" & strTxtUnAnsweredTopics & """ title=""" & strTxtUnAnsweredTopics & """ style=""vertical-align: text-bottom"" /> <a href=""active_topics.asp?UA=Y" & strQsSID2 & """>" & strTxtUnAnsweredTopics & "</a> "
'If RSS XML enabled then display an RSS button to link to XML file
If blnRSS Then strStatusBarTools = strStatusBarTools & "&nbsp;<a href=""RSS_topic_feed.asp" & SeoUrlTitle(strMainForumName, "?title=") & """ target=""_blank""><img src=""" & strImagePath & "rss." & strForumImageType & """ alt=""" & strTxtRSS & ": " & strTxtNewPostFeed & """ title=""" & strTxtRSS & " - " & strTxtNewPostFeed & """ /></a>"

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title><% = strMainForumName %></title>
<meta name="generator" content="Web Wiz Forums <% = strVersion %>" />
<meta name="description" content="<% = strBoardMetaDescription %>" />
<meta name="keywords" content="<% = strBoardMetaKeywords %>" />
<link rel="canonical" href="<% = strForumPath %>" /><%


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
If blnRSS Then Response.Write(vbCrLf & "<link rel=""alternate"" type=""application/rss+xml"" title=""RSS 2.0"" href=""RSS_topic_feed.asp" & SeoUrlTitle(strMainForumName, "?title=") & """ />")

%>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="1" cellpadding="3" align="center">
 <tr>
  <td class="smText"><% = strTxtTheTimeNowIs & " " & TimeFormat(now()) %><br /><%

'If this is not the first time the user has visted the site display the last visit time and date
If IsDate(getCookie("lVisit", "LV")) Then
	If dtmLastVisitDate < DateC(getCookie("lVisit", "LV")) Then
	   	Response.Write(strTxtYouLastVisitedOn & " " & DateFormat(dtmLastVisitDate) & " " & strTxtAt & " " & TimeFormat(dtmLastVisitDate))
	End If
End If

%><br /></td><%

Response.Write(vbCrLf & " </tr>")


 %>
</table>
<br /><%



'Check there are categories to display
If lngTotalRecords = 0 Then
	
%>
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
 <tr class="tableLedger">
  <td width="5%">&nbsp;</td>
  <td width="50%"><% = strTxtForum %></td>
  <td width="10%" align="center"><% = strTxtTopics %></td>
  <td width="10%" align="center"><% = strTxtPosts %></td>
  <td width="30%" align="center"><% = strTxtLastPost %></td>
 </tr>
 <tr>
  <td colspan="5" class="tableRow"><% = strTxtNoForums %></td>
 </tr>
</table>
<br /><%


'Else there the are categories so write the HTML to display categories and the forum names and a discription
Else	

	'Loop round to show all the categories and forums
	Do While intCurrentRecord <= Ubound(sarryForums,2)
	
		
		'Loop through the array looking for forums that are to be shown
		'if a forum is found to be displayed then show the category and the forum, if not the category is not displayed as there are no forums the user can access
		Do While intCurrentRecord <= Ubound(sarryForums,2)
		
			'Read in details
			blnHideForum = CBool(sarryForums(13,intCurrentRecord))
			blnRead = CBool(sarryForums(14,intCurrentRecord))
					
			'If this forum is to be shown then leave the loop and display the cat and the forums
			If blnHideForum = False OR blnRead Then Exit Do
			
			'Move to next record
			intCurrentRecord = intCurrentRecord + 1
		Loop
				
		'If we have run out of records jump out of loop
		If intCurrentRecord > Ubound(sarryForums,2) Then Exit Do
	
	
		'Read in the details from the array of this category
		intCatID = CInt(sarryForums(0,intCurrentRecord))
		strCategory = sarryForums(1,intCurrentRecord)		
		

%>
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center"><%

		'Display column headings if not a mobile
		If blnMobileBrowser = False Then

%>
 <tr class="tableLedger">
  <td width="5%">&nbsp;</td>
  <td width="50%"><% = strTxtForum %></td>
  <td width="10%" align="center"><% = strTxtTopics %></td>
  <td width="10%" align="center"><% = strTxtPosts %></td>
  <td width="30%"><% = strTxtLastPost %></td>
 </tr><%
		
		End If
		
		'Display the category name
		Response.Write vbCrLf & " <tr class=""tableSubLedger""><td colspan=""5""><a href=""default.asp?C=" & intCatID & strQsSID2 & SeoUrlTitle(strCategory, "&title=") & """>" & strCategory & "</a></td></tr>"
		
		'If the user only wants to see one category, only display the forums for that category
		If intCatShow = intCatID OR intCatShow = 0 Then
			
			'Loop round to display all the forums for this category
			Do While intCurrentRecord <= Ubound(sarryForums,2)
				
				'Initialise variables
				strSubForums = ""
			
			
				'Read in the details for this forum
				intForumID = CInt(sarryForums(2, intCurrentRecord))
				intSubForumID = CInt(sarryForums(3, intCurrentRecord))
				strForumName = sarryForums(4, intCurrentRecord)
				strForumDiscription = sarryForums(5, intCurrentRecord)
				lngNumberOfTopics = CLng(sarryForums(6, intCurrentRecord))
				lngNumberOfPosts = CLng(sarryForums(7, intCurrentRecord))
				strLastEntryUser = sarryForums(8, intCurrentRecord)
				If isNumeric(sarryForums(9, intCurrentRecord)) Then lngLastEntryUserID = CLng(sarryForums(9, intCurrentRecord)) Else lngLastEntryUserID = 0
				If isDate(sarryForums(10, intCurrentRecord)) Then dtmLastEntryDate = CDate(sarryForums(10, intCurrentRecord)) Else dtmLastEntryDate = CDate("2001-01-01 00:00:00")
				If isNull(sarryForums(11, intCurrentRecord)) Then strForumPassword = "" Else strForumPassword = sarryForums(11, intCurrentRecord)
				blnForumLocked = CBool(sarryForums(12, intCurrentRecord))
				blnHideForum = CBool(sarryForums(13, intCurrentRecord))
				blnRead = CBool(sarryForums(14, intCurrentRecord))
				If isNumeric(sarryForums(15, intCurrentRecord)) Then lngTopicID = CLng(sarryForums(15, intCurrentRecord)) Else lngTopicID = 0
				strSubject = sarryForums(16, intCurrentRecord)	
				strForumImageIcon = sarryForums(17, intCurrentRecord)	
				strForumURL = sarryForums(18, intCurrentRecord)	
				
				'Set the last forum ID for forum stats
				intLastForumEntryID = intForumID
				
				'Remove any parts that could be mistaken for a forum URL
				If strForumURL = "http://" OR isNull(strForumURL) Then strForumURL = ""

				'If this forum is to be hidden but the user is allowed access to it set the hidden boolen back to false
				If blnHideForum AND blnRead Then blnHideForum = False

				'Calculate the number of people viewing the forum
				If blnForumViewing AND blnActiveUsers Then 
					intTotalViewingForum = viewingForum(intForumID)
				End If
				
				'If the forum is not a hidden forum to this user, display it
				If blnHideForum = False AND intSubForumID = 0 Then

					
					'Stats ***********
					'Count the number of forums
					intNumberofForums = intNumberofForums + 1
	
					'Add all the posts and topics together to get the total number for the stats at the bottom of the page
					lngTotalNumberOfPosts = lngTotalNumberOfPosts + lngNumberOfPosts
					lngTotalNumberOfTopics = lngTotalNumberOfTopics + lngNumberOfTopics
					
					'Calculate the last forum entry across all forums for the statistics at the bottom of the forum
					If dtmLastEntryDateAllForums < dtmLastEntryDate Then
						strLastEntryUserAllForums = strLastEntryUser
						lngLastEntryUserIDAllForums = lngLastEntryUserID
						dtmLastEntryDateAllForums = dtmLastEntryDate
					End If
					
					
					
					
					'Unread Posts *********
					intUnReadPostCount = 0
					
					'If there is a newer post than the last time the unread posts array was initilised run it again
					If dtmLastEntryDate > CDate(Session("dtmUnReadPostCheck")) Then Call UnreadPosts()
						
					'Count the number of unread posts in this forum
					If isArray(sarryUnReadPosts) AND dtmLastEntryDate > dtmLastVisitDate AND blnRead Then
						For intUnReadForumPostsLoop = 0 to UBound(sarryUnReadPosts,2)
							'Increament unread post count
							If CInt(sarryUnReadPosts(2,intUnReadForumPostsLoop)) = intForumID AND sarryUnReadPosts(3,intUnReadForumPostsLoop) = "1" Then intUnReadPostCount = intUnReadPostCount + 1
						Next
						
						'Get the text for unread post
						If intUnReadPostCount = 1 Then strNewPostText = strTxtNewPost Else strNewPostText = strTxtNewPosts
					End If
					
					
					
					
					'Get the row number
					intForumColourNumber = intForumColourNumber + 1


					
					'Display if this forum has any subforums
					'***************************************
					
					'Initilise variables
					intTempRecord = 0
					
					'Loop round to read in any sub forums in the stored array recordset
					Do While intTempRecord <= Ubound(sarryForums,2) 
					
					
						'Becuase the member may have an individual permission entry in the permissions table for this forum, 
						'it maybe listed twice in the array, so we need to make sure we don't display the same forum twice
						If intSubForumID = CInt(sarryForums(2,intTempRecord)) Then intTempRecord = intTempRecord + 1
							
						'If there are no records left exit loop
						If intTempRecord > Ubound(sarryForums,2) Then Exit Do
				
						'If this is a subforum of the main forum then get the details
						If CInt(sarryForums(3,intTempRecord)) = intForumID Then
													
						
							'Read in sub forum details from the database
							intSubForumID = CInt(sarryForums(2,intTempRecord))
							strSubForumName = sarryForums(4,intTempRecord)
							lngSubForumNumberOfTopics = CLng(sarryForums(6,intTempRecord))
							lngSubForumNumberOfPosts = CLng(sarryForums(7,intTempRecord))
							strLastSubEntryUser = sarryForums(8,intTempRecord)
							If isNumeric(sarryForums(9,intTempRecord)) Then lngLastSubEntryUserID = CLng(sarryForums(9,intTempRecord)) Else lngLastEntryUserID = 0
							If isDate(sarryForums(10,intTempRecord)) Then dtmLastSubEntryDate = CDate(sarryForums(10,intTempRecord)) Else dtmLastSubEntryDate = CDate("2001-01-01 00:00:00")
							If isNull(sarryForums(11, intCurrentRecord)) Then strSubForumPassword = "" Else strSubForumPassword = sarryForums(11, intCurrentRecord)
							blnHideForum = CBool(sarryForums(13,intTempRecord))	
							blnSubRead = CBool(sarryForums(14,intTempRecord))
							If isNumeric(sarryForums(15, intTempRecord)) Then lngSubTopicID = CLng(sarryForums(15, intTempRecord)) Else lngSubTopicID = 0
							
							strSubSubject = sarryForums(16, intTempRecord)	
						
				
							'If this sub forum is to be hidden and but the user is allowed access to it set the hidden boolen back to false
							If blnHideForum = True AND blnSubRead = True Then blnHideForum = False
						
							'If the sub forum is to be hidden then don't show it
							If blnHideForum = False Then
								
								'Stats **********
								'Count the number of forums
								intNumberofForums = intNumberofForums + 1
								
								'Add all the posts and topics together to get the total number for the stats at the bottom of the page
								lngTotalNumberOfPosts = lngTotalNumberOfPosts + lngSubForumNumberOfPosts
								lngTotalNumberOfTopics = lngTotalNumberOfTopics + lngSubForumNumberOfTopics
								
								'Add the number of posts and topics of sub-forums to the number of posts in the main forum
								lngNumberOfPosts = lngNumberOfPosts + lngSubForumNumberOfPosts
								lngNumberOfTopics = lngNumberOfTopics + lngSubForumNumberOfTopics
								
								
								'Calculate the last forum entry across all forums for the statistics at the bottom of the forum
								If dtmLastEntryDateAllForums < dtmLastSubEntryDate Then
									strLastEntryUserAllForums = strLastSubEntryUser
									lngLastEntryUserIDAllForums = lngLastSubEntryUserID
									dtmLastEntryDateAllForums = dtmLastSubEntryDate
								End If
								
								'If the subforums last entry is newer than that of the main forum, then display that as the last post in the forum
								If (dtmLastEntryDate < dtmLastSubEntryDate) AND blnSubRead AND strSubForumPassword = "" Then
									intLastForumEntryID = intSubForumID
									strLastEntryUser = strLastSubEntryUser
									lngLastEntryUserID = lngLastSubEntryUserID
									dtmLastEntryDate = dtmLastSubEntryDate
									lngTopicID = lngSubTopicID
									strSubject = strSubSubject
								End If
								
								'Unread Posts *********
								
								'If there is a newer post than the last time the unread posts array was initilised run it again
								If dtmLastSubEntryDate > CDate(Session("dtmUnReadPostCheck")) Then Call UnreadPosts()
									
								'Count the number of unread posts in this forum
								If isArray(sarryUnReadPosts) AND dtmLastSubEntryDate > dtmLastVisitDate AND blnSubRead Then
									For intUnReadForumPostsLoop = 0 to UBound(sarryUnReadPosts,2)
										'Increament unread post count
										If CInt(sarryUnReadPosts(2,intUnReadForumPostsLoop)) = intSubForumID AND sarryUnReadPosts(3,intUnReadForumPostsLoop) = "1" Then intUnReadPostCount = intUnReadPostCount + 1
									Next	
								End If
								
								'Calculate the number of people viewing the forum
								If blnForumViewing AND blnActiveUsers Then 
									intTotalViewingForum = intTotalViewingForum + viewingForum(intSubForumID)
								End If
								
								'If there are other sub forums place a comma inbetween
								If strSubForums <> "" Then strSubForums = strSubForums & ", "
								
								'Display the sub forum
								If blnUrlRewrite Then
									strSubForums = strSubForums & "<a href=""" & SeoUrlTitle(strSubForumName, "") & "_forum" & intSubForumID & ".html" & strQsSID1 & """ class=""smLink"">" & strSubForumName & "</a>"
								Else
									strSubForums = strSubForums & "<a href=""forum_topics.asp?FID=" & intSubForumID & strQsSID2 & SeoUrlTitle(strSubForumName, "&title=") & """ class=""smLink"">" & strSubForumName & "</a>"
								End If
							End If
						End If
						
						'Move to next record 
						intTempRecord = intTempRecord + 1
					Loop
					
					
					
					'If mobile browser display different content
					If blnMobileBrowser Then
						
						'Calculate row colour
						Response.Write(vbCrLf & " <tr ")
						If (intForumColourNumber MOD 2 = 0 ) Then Response.Write("class=""evenTableRow"">") Else Response.Write("class=""oddTableRow"">") 
						
						'Display link to forum
						If strForumURL = "" Then
							Response.Write("<td>")
							
							If blnUrlRewrite Then
								Response.Write("<a href=""" & SeoUrlTitle(strForumName, "") & "_forum" & intForumID & ".html" & strQsSID1 & """>" & strForumName & "</a>")
							Else
								Response.Write("<a href=""forum_topics.asp?FID=" & intForumID & strQsSID2 & SeoUrlTitle(strForumName, "&title=") & """>" & strForumName & "</a>")
							End If
							
							'Display unread post count to mobile users
							If intUnReadPostCount = 1 Then
								Response.Write(" [1 " & strTxtNewPost & "]")
							ElseIf intUnReadPostCount > 1 Then
								Response.Write(" [" & intUnReadPostCount & " " & strTxtNewPosts & "]")
							End If
							
							'Display sub forums to mobile users
							'If strSubForums <> "" Then Response.Write("<br /><span class=""smText"">" & strTxtSub & " " & strTxtForums & ": </span>" & strSubForums)
							
							Response.Write("</td></tr>")	
						
						'Else is a link
						Else
							Response.Write("<td><a href=""" & strForumURL & """>" & strForumName & "</a></td></tr>")
						End If
					
					
					
					'Else not a mobile browser so display normal tables
					Else

						'If there are sub forums 
						If strSubForums <> "" Then strSubForums = "<br /><span class=""smText"">" & strTxtSub & " " & strTxtForums & ": </span>" & strSubForums
	
	
						'Write the HTML of the forum descriptions and hyperlinks to the forums
						
						'Calculate row colour
						Response.Write(vbCrLf & " <tr ")
						If (intForumColourNumber MOD 2 = 0 ) Then Response.Write("class=""evenTableRow"">") Else Response.Write("class=""oddTableRow"">") 
			
						'If not a external link
						If strForumURL = "" Then
							
							'Display the status forum icons
							Response.Write(vbCrLf & "   <td align=""center"">")
							%><!-- #include file="includes/forum_status_icons_inc.asp" --><%
		     					Response.Write("</td>" & _
							vbCrLf & "  <td>")
		
		
							
							'Display forum
							If blnUrlRewrite Then
								Response.Write("<a href=""" & SeoUrlTitle(strForumName, "") & "_forum" & intForumID & ".html" & strQsSID1 & """>" & strForumName & "</a>")
							Else
								Response.Write("<a href=""forum_topics.asp?FID=" & intForumID & strQsSID2 & SeoUrlTitle(strForumName, "&title=") & """>" & strForumName & "</a>")
							End If
							
							
					
							'Display the number of people viewing in that forum
							If blnForumViewing AND blnActiveUsers Then 
								If intTotalViewingForum > 0 Then Response.Write(" <span class=""smText"">(" & intTotalViewingForum & " " & strTxtViewing & ")</span>")
							End If
							
							'Display forum details
							Response.Write("<br />" & strForumDiscription & strSubForums & "</td>" & _
							vbCrLf & "  <td align=""center"">" & lngNumberOfTopics & "</td>" & _
							vbCrLf & "  <td align=""center"">" & lngNumberOfPosts & "</td>" & _
							vbCrLf & "  <td class=""smText"" nowrap=""nowrap"">")
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
									strSubject = formatInput(strSubject)
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
										Response.Write("<a href=""get_last_post.asp?TID=" & lngTopicID & "&amp;ID=" & dblGetLastPostID & strQsSID2 & """><img src=""" & strImagePath & "view_unread_post." & strForumImageType & """ alt=""" & strTxtViewUnreadPost & " [" & intUnReadPostCount & " " & strNewPostText & "]"" title=""" & strTxtViewUnreadPost & " [" & intUnReadPostCount & " " & strNewPostText & "]"" /></a>")
									
									'Else there are no unread posts so display a normal last post link
									Else
										Response.Write("<a href=""get_last_post.asp?TID=" & lngTopicID & "&amp;ID=" & dblGetLastPostID & strQsSID2 & """><img src=""" & strImagePath & "view_last_post." & strForumImageType & """ alt=""" & strTxtViewLastPost & """ title=""" & strTxtViewLastPost & """ /></a>")
									End If
									
								End If
								'Last Post date
								Response.Write("<br />" & DateFormat(dtmLastEntryDate) & "&nbsp;" &  strTxtAt & "&nbsp;" & TimeFormat(dtmLastEntryDate))
							End If	
							Response.Write("</td>"  & _
							vbCrLf & " </tr>")
						
						'Else if forum link
						Else
							
							'Display extrenal link row
							Response.Write(vbCrLf & "   <td align=""center"">")
							
							'Display a custom icon is used for the forum
							If NOT strForumImageIcon = "" Then  
								Response.Write("<img src=""" & strForumImageIcon & """ border=""0"" alt=""" & strForumIconTitle & """ title=""" & strForumIconTitle & """ />")	
							Else
								Response.Write("<img src=""" & strImagePath & "web_link." & strForumImageType & """ border=""0"" alt=""" & strTxtExternalLinkTo & ": " & strForumURL & """ title=""" & strTxtExternalLinkTo & ": " & strForumURL & """ />")	
							End If
							
							'Display extrenal link
		     					Response.Write("</td><td colspan=""4""><a href=""" & strForumURL & """>" & strForumName & "</a><br />" & strForumDiscription & strSubForums & "</td></td></tr>")
							
						End If
						
					End If

				End If

				

				'Move to the next database record
				intCurrentRecord = intCurrentRecord + 1
				
				
				'If there are more records in the array to display then run some test to see what record to display next and where				
				If intCurrentRecord <= Ubound(sarryForums,2) Then

					'Becuase the member may have an individual permission entry in the permissions table for this forum, 
					'it maybe listed twice in the array, so we need to make sure we don't display the same forum twice
					If intForumID = CInt(sarryForums(2,intCurrentRecord)) Then intCurrentRecord = intCurrentRecord + 1
					
					'If there are no records left exit loop
					If intCurrentRecord > Ubound(sarryForums,2) Then Exit Do
					
					'If this is a subforum jump to the next record, unless we have run out of forums
					Do While CInt(sarryForums(3,intCurrentRecord)) > 0 
						
						'Go to next record
						intCurrentRecord = intCurrentRecord + 1
						
						'If we have run out of records jump out of loop
						If intCurrentRecord > Ubound(sarryForums,2) Then Exit Do
					Loop
					
					'If there are no records left exit loop
					If intCurrentRecord > Ubound(sarryForums,2) Then Exit Do
					
					'See if the next forum is in a new category, if so jump out of this loop to display the next category
					If intCatID <> CInt(sarryForums(0,intCurrentRecord)) Then Exit Do
				End If
			
			'Loop back round to display next forum
			Loop
		
		
		'Else we are not displaying forums in this category so we need to move to the next category in the array
		Else
		
			'Loop through the forums array till we get to the next category
			Do While CInt(sarryForums(0,intCurrentRecord)) = intCatID 
					
				'Go to next record
				intCurrentRecord = intCurrentRecord + 1
					
				'If we have run out of records jump out of forums loop into the category loop
				If intCurrentRecord > Ubound(sarryForums,2) Then Exit Do
			Loop	
		End If

		
	
	%>
</table>
<br /><%
     
	'Loop back round for next category
	Loop
End If



'Display list of latest forum posts
If blnShowLatestPosts Then
	%><!--#include file="includes/latest_posts_inc.asp" --><%
End If

'Clean up
Call closeDatabase()

'If NewsPad is enabled and we have a URL to it display the Web Wiz NewsPad new bulletins
If blnWebWizNewsPad AND strWebWizNewsPadURL <> "" Then
%>

<span id="WebWizNewsPad"></span>
<script language="javascript" type="text/javascript">getAjaxData('ajax_newspad_feed.asp', 'WebWizNewsPad');</script>
<br /><%

End If


'Do not display stats for mobile browsers
If blnMobileBrowser = False AND (blnDisplayForumStats OR blnActiveUsers) Then
%>
<br />
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
 <tr class="tableLedger">
  <td colspan="2"><% = strTxtWhatsGoingOn %></td>
 </tr><%
 
 	'Display forum statis if enabled
 	If blnDisplayForumStats Then

%>
 <tr class="tableSubLedger">
  <td colspan="2"><% = strTxtForumStatistics %></td>
 </tr>
 <tr class="tableRow">
  <td width="5%" align="center"><img src="<% = strImagePath %>forum_statistics.<% = strForumImageType %>" alt="<% = strTxtForumStatistics %>" title="<% = strTxtForumStatistics %>" /></td>
  <td width="95%" nowrap="nowrap"><%

	Response.Write(strTxtOurUserHavePosted & " " & FormatNumber(lngTotalNumberOfPosts, 0) & " " & strTxtPostsIn & " " & FormatNumber(lngTotalNumberOfTopics, 0) & " " & strTxtTopicsIn & " " & intNumberofForums & " " & strTxtForums & _
	"<br />" & strTxtLastPost & ", " & DateFormat(dtmLastEntryDateAllForums) & " " & strTxtAt & " " & TimeFormat(dtmLastEntryDateAllForums) & " " & strTxtBy & " <a href=""member_profile.asp?PF=" & lngLastEntryUserIDAllForums & strQsSID2 & """>" & strLastEntryUserAllForums & "</a>")
	
	'Display some statistics for the members
	If lngNoOfMembers > 0 Then
	
	Response.Write("<br />" & strTxtWeHave & " " & FormatNumber(lngNoOfMembers, 0) & " " & strTxtForumMembers & _
	"<br />" & strTxtTheNewestForumMember & " <a href=""member_profile.asp?PF=" & saryMemebrStats(1,0) & strQsSID2 & """>" & saryMemebrStats(0, 0) & "</a>")
	
	End If

%></td>
 </tr><%
 
	End If


	'Get the number of active users if enabled
	If blnActiveUsers Then
	
	%>
 <tr class="tableSubLedger">
  <td colspan="2"><a href="active_users.asp<% = strQsSID1 %>"><% = strTxtActiveUsers %></a></td>
 </tr>
 <tr class="tableRow">
  <td width="5%" align="center"><a href="active_users.asp<% = strQsSID1 %>"><img src="<% = strImagePath %>active_users.<% = strForumImageType %>" alt="<% = strTxtActiveUsers %>" title="<% = strTxtView & " " & strTxtActiveUsers %>" border="0" /></a></td>
  <td width="95%"><%
	
		'Get the active users online
		For intArrayPass = 1 To UBound(saryActiveUsers, 2)
		
			'If this is a search robot then increment serach robot
			If Instr(saryActiveUsers(5, intArrayPass), "Search Robot") Then
				
				intActiveSearchRobots = intActiveSearchRobots + 1
				
			'If this is a guest user then increment the number of active guests veriable	
			ElseIf saryActiveUsers(1, intArrayPass) = 2 Then 
					
				intActiveGuests = intActiveGuests + 1
			
			'Else if the user is Anonymous increment the Anonymous count
			ElseIf CBool(saryActiveUsers(8, intArrayPass)) Then	
				
				intAnonymousMembers = intAnonymousMembers + 1
			
			'Else add the name of the members name of the active users to the members online string
			ElseIf CBool(saryActiveUsers(8, intArrayPass)) = false Then	
				If strMembersOnline <> "" Then strMembersOnline = strMembersOnline & ", "
				strMembersOnline = strMembersOnline & "<a href=""member_profile.asp?PF=" & saryActiveUsers(1, intArrayPass) & strQsSID2 & """>" & saryActiveUsers(2, intArrayPass) & "</a>"
			End If
			
		Next 
	
		'Calculate the number of members online and total people online
		intActiveUsers = UBound(saryActiveUsers, 2)
		
		'Calculate the members online by using the total - Guests - Annoymouse Members
		intActiveMembers = intActiveUsers - intActiveGuests - intAnonymousMembers - intActiveSearchRobots
	
		
		Response.Write(strTxtInTotalThereAre & " " & intActiveUsers & " <a href=""active_users.asp" & strQsSID1 & """>" & strTxtActiveUsers & "</a> " & strTxtOnLine & ", " & intActiveMembers & " " & strTxtMembers & ", " & intAnonymousMembers & " " & strTxtAnonymousMembers & ", " & intActiveGuests & " " & strTxtGuests & ", " & intActiveSearchRobots & " " & strTxtSearchRobots & _
			vbCrLf & "   <br />" & strTxtMostUsersEverOnlineWas & " " & lngMostEverActiveUsers & ", " & DateFormat(dtmMostEverActiveDate) & " " & strTxtAt & " " & TimeFormat(dtmMostEverActiveDate))
		If strMembersOnline <> "" Then Response.Write(vbCrLf & "   <br />" & strTxtMembers & " " & strTxtOnLine & ": " & strMembersOnline)
	End If

%>
  </td>
 </tr><%
 

	'If birthdays is enabled show who has a birthday today
	If strBirthdays <> "" Then
 
 %>
 <tr class="tableSubLedger">
  <td colspan="2"><% = strTxtTodaysBirthdays %></td>
 </tr>
 <tr class="tableRow">
  <td width="5%" align="center"><img src="<% = strImagePath %>todays_birthdays.<% = strForumImageType %>" alt="<% = strTxtTodaysBirthdays %>" title="<% = strTxtTodaysBirthdays %>" /></td>
  <td width="95%"><% = strBirthdays %></td>
 </tr><%
 
	End If
 
 %>
</table><%

End If

%>
<br />
<div align="center">
<span class="smText"><a href="mark_posts_as_read.asp<% If strSessionID <> "" Then Response.Write("?XID=" & getSessionItem("KEY") & strQsSID2) %>" class="smLink"><% = strTxtMarkAllPostsAsRead %></a> :: <a href="remove_cookies.asp<% If strSessionID <> "" Then Response.Write("?XID=" & getSessionItem("KEY") & strQsSID2) %>" class="smLink"><% = strTxtDeleteCookiesSetByThisForum %></a> :: <a href="help.asp<% = strQsSID2 %>#FAQ36" class="smLink"><% = strTxtCookiePrivacyNotice %></a>
<br /></span><%

'If a mobile browser display an option to switch to and from mobile view
If blnMobileBrowser Then 
	Response.Write ("<br />" & strTxtViewIn & ": <strong>" & strTxtMoble & "</strong> | <a href=""?MobileView=off" & strQsSID2 & """ rel=""nofollow"">" & strTxtClassic & "</a>")
ElseIf blnMobileClassicView Then
	Response.Write ("<br />" & strTxtViewIn & ": <a href=""?MobileView=on" & strQsSID2 & """ rel=""nofollow"">" & strTxtMoble & "</a> | <strong>" & strTxtClassic & "</strong>")
End If

%><span class="smText"><br /><br /><% = strTxtCookies %></span><br /><br /><%
    
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnLCode = True Then
	If blnTextLinks = True Then
		Response.Write("<span class=""text"" style=""font-size:10px"">Forum Software by <a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank""  style=""font-size:10px"">Web Wiz Forums&reg;</a> version " & strVersion & "</span>")
	Else
  		Response.Write("<a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank"" ><img src=""webwizforums_image.asp"" border=""0"" title=""Forum Software by Web Wiz Forums&reg; version " & strVersion & """ alt=""Forum Software by Web Wiz Forums&reg; version " & strVersion& """ /></a>")
	End If

	Response.Write("<br /><span class=""text"" style=""font-size:10px"">Copyright &copy;2001-2019 Web Wiz Ltd.</span>")
End If
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******

'Display the process time
If blnShowProcessTime Then Response.Write "<span class=""smText""><br /><br />" & strTxtThisPageWasGeneratedIn & " " & FormatNumber(Timer() - dblStartTime, 3) & " " & strTxtSeconds & "</span>"

%>
</div>
   <!-- #include file="includes/footer.asp" -->