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


'Declare vars
Dim intSearchInterval		'Holds the amount of time between searches
Dim intMaxResults		'Holds the maximum returned results
Dim intShowTopicsFrom		'Holds when to show topic from
Dim strSQLwhereDate		'Holds the SQL for the date to get results for
Dim iarySearchForumID		'Holds the forums to be searched
Dim intCurrentRecord		'Holds the current record position
Dim strMemberName		'Holds the member name to search for
Dim intDateDirection		'Set if we are looking before or after trhe date given
Dim strSearchKeywords		'Holds the keywords to be search for
Dim sarySearchWord		'Array holding each of the search keywords
Dim strSQLwhereKeywords		'Holds the SQL for the seach oart of the sql query
Dim strTableFieldName		'Holds the table field
Dim strSQLoperator		'Holds AND or OR
Dim strSQLwhereForum		'Holds which forum to search
Dim strSQLwhereMemSearch	'Holds which user to search for
Dim strSearchType		'Holds the search type
Dim strSearhIn			'Holds where we are searching
Dim strDateDirection		'Holds the direction to run the searc in (before or after date)
Dim strForumIDs			'Holds the forum ID to search in
Dim blnExactUserMatch		'Set to true if the username needs to be exact
Dim strOrderBy			'Holds the order by cluse for the sql query
Dim sarySearchIndex		'Holds the details of the search array
Dim strSeachIPAddress		'Holds the user IP
Dim blnSearchIPFound 		'Set to true if the IP address has run a search recently
Dim lngSearchID			'Holds the search ID
Dim intSearchTimeToLive		'Holds the time to keep the search in memory
Dim blnResultsFound		'Set to true of no results found
Dim intRemovedEntries		'Holds the removed entries
Dim blnSearhWordsTwoShort	'Set to true if search words are below 4 chars
Dim intSearchWordLength		'Holds the length of the search words
Dim strResultType		'Holds the result type of how the results are displayed
Dim lngTopicID			'Holds the topic id if doing a topic search
Dim blnMemberStarted		'Set to true if we are looking for posts a member has started
Dim blnTimeoutError




'Initialise variables
intCurrentRecord = 0
intRemovedEntries = 0
blnSearchIPFound = false
blnResultsFound = false
blnSearhWordsTwoShort = false
blnTimeoutError = false



'Test querystrings for any SQL Injection keywords
Call SqlInjectionTest(Request.QueryString())





'******************************************
'***   Search Settings			***
'******************************************

'Search process ASP timeout
Server.ScriptTimeout =  120 'Amount of secounds this process can run

'Search database connection timemout (this maybe overridden by settings on the datbase server)
adoCon.CommandTimeout = 60	'Amount of secounds before the database connection timesout

'Time to keep the search results in memory
intSearchTimeToLive = 20 	'Amount of minutes to keep the search in memory	


If strResultType = "topics" Then
	intMaxResults = 100 'The max amount of results when searching as topics (topics requires more joins so is less efficeint)
Else
	intMaxResults = 100 'The max amount of results when search as posts
End If


'Set the time between searches
'If not logged in then 30 secounds
If intGroupID = 2 Then
	intSearchInterval = 30 'Length of time between new searches for non-logged in users
	intSearchWordLength = 2 'Holds the minimum search word length non-logged in users
'Else logged in users 0 secounds
Else
	intSearchInterval = 0
	intSearchWordLength = 1 'Holds the minimum search word length for loged in users
End IF






'Get the users IP address
strSeachIPAddress = getIP()


'******************************************
'***   	Read in form details		***
'******************************************

'Read in form input
strSearchKeywords = Trim(Mid(Request.Form("KW"), 1, 35))
strSearchType = Request.Form("searchType")
strMemberName = Trim(Mid(Request.Form("USR"), 1, 20))
blnExactUserMatch = BoolC(Request.Form("UsrMatch"))
blnMemberStarted = BoolC(Request.Form("UsrTopicStart"))
strSearhIn = Request.Form("searchIn")
intShowTopicsFrom = IntC(Request.Form("AGE"))
strDateDirection = Request.Form("DIR")
strForumIDs = Request.Form("forumID")
strOrderBy = Request.Form("OrderBy")
strResultType = Request.Form("resultType")


'Set up topic search, if searching in a topic
If Request.Form("qTopic") Then 
	lngTopicID = LngC(Request.Form("TID"))
	strSearhIn = "body"
	intShowTopicsFrom = 0
	strOrderBy = "StartDate"
	strResultType = "posts"
End If	






'******************************************
'***   	Initialise  array		***
'******************************************
	
'Initialise  the array from the application veriable
If IsArray(Application(strAppPrefix & "sarySearchIndex")) Then 
	
	'Place the application level search results array into a temporary dynaimic array
	sarySearchIndex = Application(strAppPrefix & "sarySearchIndex")


'Else Initialise the an empty array
Else
	ReDim sarySearchIndex(5,0)
End If


'Array dimension lookup table
' 0 = Search ID
' 1 = IP
' 2 = Date/time last run
' 3 = Member ID
' 4 = Date/time search created
' 5 = time taken to run search



'**************************************************
'***  IP check to limit no. of searches		***
'**************************************************

'Iterate through the array to see if the IP address has ran a search in the last xx secounds
For intCurrentRecord = 0 To UBound(sarySearchIndex, 2)
	If sarySearchIndex(1, intCurrentRecord) = strSeachIPAddress AND CDate(sarySearchIndex(4, intCurrentRecord)) > DateAdd("s", - intSearchInterval, Now()) Then 
		blnSearchIPFound = true
	End If
Next 






'******************************************
'***   	Remove old search results	***
'******************************************

'Iterate through the array to see if the search is already in memory and see if the IP address has ran a search in the last xx secounds
For intCurrentRecord = 1 To UBound(sarySearchIndex, 2)
	
	'Check the serach results are not old, if they are remove them
	If CDate(sarySearchIndex(2, intCurrentRecord)) < DateAdd("n", - intSearchTimeToLive, Now()) Then
		
		'Delete the search result application array
		'Lock the application so that no other user can try and update the application level variable at the same time
		Application.Lock
		
		'Distroy the application level array
		Application(sarySearchIndex(0, intCurrentRecord)) = null
		
		'Unlock the application
		Application.UnLock
		
		'Swap this array postion with the last in the array
		sarySearchIndex(0, intCurrentRecord) = sarySearchIndex(0, UBound(sarySearchIndex, 2))
		sarySearchIndex(1, intCurrentRecord) = sarySearchIndex(1, UBound(sarySearchIndex, 2))
		sarySearchIndex(2, intCurrentRecord) = sarySearchIndex(2, UBound(sarySearchIndex, 2))
		sarySearchIndex(3, intCurrentRecord) = sarySearchIndex(3, UBound(sarySearchIndex, 2))
		sarySearchIndex(4, intCurrentRecord) = sarySearchIndex(4, UBound(sarySearchIndex, 2))
		sarySearchIndex(5, intCurrentRecord) = sarySearchIndex(5, UBound(sarySearchIndex, 2))
		
		'Set how many entries to remove
		intRemovedEntries = intRemovedEntries + 1
	End If
Next

'Remove the end array entries that are no-longer needed
If intRemovedEntries <> 0 Then ReDim Preserve sarySearchIndex(5, UBound(sarySearchIndex, 2) - intRemovedEntries)


'Reset current record variable
intCurrentRecord = 0



'Remove mutiple spaces
strSearchKeywords = Replace(strSearchKeywords, "    ", " ")
strSearchKeywords = Replace(strSearchKeywords, "   ", " ")
strSearchKeywords = Replace(strSearchKeywords, "  ", " ")

'Filter for SQL injections
strSearchKeywords = formatSQLInput(strSearchKeywords)


'If there is nothing to seardh for don't run search
If strSearchKeywords = "" AND strMemberName = "" Then blnSearhWordsTwoShort = True


'******************************************
'***   	Build SQL Query for Search	***
'******************************************

If blnSearchIPFound = False AND blnSearhWordsTwoShort = False Then


	'******************************************
	'***   	SQL for keyword search		***
	'******************************************
	 
	'Build the SQL search string
	If strSearchKeywords <> "" Then
		
		'If searcing in a topic
		If strSearhIn = "subject" Then
			
			'Filter more if a topic subject search, as topic subjects are filtered more
			strSearchKeywords = removeAllTags(strSearchKeywords)
	
			'Set the field name for the SQL query
			strTableFieldName = strDbTable & "Topic.Subject"
		
		
		'Else this is a search of the post
		Else
			'Set the field name for the SQL query
			strTableFieldName = strDbTable & "Thread.Message"
			
			
			'If displaying results in topic view use a sub query
			If strResultType = "topics" Then
				
				strSQLwhereKeywords = strSQLwhereKeywords & "" & _
				"AND (" & strDbTable & "Topic.Topic_ID " & _
					"IN (" & _
						"SELECT " & strDbTable & "Thread.Topic_ID " & _
						"FROM " & strDbTable & "Thread" & strDBNoLock & " " & _
						"WHERE " 
			End If
			
		End If
		
		
		
		'Create the SQL
		'If searching for an IP address
		If strSearchType = "IP" Then 
			If strResultType = "topics" AND strSearhIn = "body" Then 
				strSQLwhereKeywords = strSQLwhereKeywords & " (" & strDbTable & "Thread.IP_addr LIKE '%" & strSearchKeywords & "%') "
			Else
				strSQLwhereKeywords = strSQLwhereKeywords & " AND (" & strDbTable & "Thread.IP_addr LIKE '%" & strSearchKeywords & "%') "
			End If
			
		
		'If this is a phrase search then check for the phrase
		ElseIf strSearchType = "phrase" Then
			'If searching in topic view then don't use AND as it is a sub query
			If strResultType = "topics" AND strSearhIn = "body" Then 
				strSQLwhereKeywords = strSQLwhereKeywords & " (" & strTableFieldName & " LIKE '%" & strSearchKeywords & "%')"
			Else
				strSQLwhereKeywords = strSQLwhereKeywords & "AND (" & strTableFieldName & " LIKE '%" & strSearchKeywords & "%')"
			End If
			'Check length
			If Len(strSearchKeywords) <= intSearchWordLength Then blnSearhWordsTwoShort = True
		
		'Else this is a Any Words or All Words search
		Else
		
			'If this is a search for Any Words use 'OR' for SQL
			If strSearchType = "anyWords" Then
				strSQLoperator = "OR"
			'Else if this is a search of All Words use 'AND' for the SQL
			Else
				strSQLoperator = "AND"
			End If
			
			'Split the search keywords and place into an array
			sarySearchWord = Split(Trim(strSearchKeywords), " ")
			
			'Build the SQL search query
			'If displaying results as topics then don't use AND
			If strResultType = "topics" AND strSearhIn = "body" Then
				strSQLwhereKeywords = strSQLwhereKeywords & " ("
			Else
				strSQLwhereKeywords = strSQLwhereKeywords & "AND ("
			End If
			
			'Loop through all the selected forums
			For intCurrentRecord = 0 To UBound(sarySearchWord)
			
				
				'If this is 2nd or more time around add OR
				If intCurrentRecord > 0 Then strSQLwhereKeywords = strSQLwhereKeywords & " " & strSQLoperator & " " 
				
				'Add the keyword to look in to the SQL query
				strSQLwhereKeywords = strSQLwhereKeywords & strTableFieldName & " LIKE '%" & sarySearchWord(intCurrentRecord) & "%'"
			
				'Check length of keywords
				If Len(sarySearchWord(intCurrentRecord)) <= intSearchWordLength Then blnSearhWordsTwoShort = True
			Next
			
			strSQLwhereKeywords = strSQLwhereKeywords & ") "
			
			'Reset record count
			intCurrentRecord = 0	
		End If
		
		'If displaying results in topic view then check if the message is hidden and close sub query
		If strResultType = "topics" AND strSearhIn = "body" Then
			If blnModerator = false AND blnAdmin = false Then 
				strSQLwhereKeywords = strSQLwhereKeywords & "AND (" & strDbTable & "Thread.Hide=" & strDBFalse & " "
				If intGroupID <> 2 Then strSQLwhereKeywords = strSQLwhereKeywords & "OR " & strDbTable & "Thread.Author_ID=" & lngLoggedInUserID
				strSQLwhereKeywords = strSQLwhereKeywords & ") "
			End If
			strSQLwhereKeywords = strSQLwhereKeywords & _
				")" & _
			") "
		End If
	End If
	
	
	
	
	'******************************************
	'***   	SQL for member search		***
	'******************************************
	
	'SQL for member search
	If strMemberName <> "" Then
		
		'Get rid of milisous code
		strMemberName = formatSQLInput(strMemberName)
		
		'Check length of member name
		If Len(strMemberName) <= intSearchWordLength Then blnSearhWordsTwoShort = True
		
		'If displaying results in topic view use a sub query
		If strResultType = "topics" Then
			
			'Create SQL for member search, using a sub query so we can get all the topics the member has posted in
			'Get only topics this meber started
			If blnMemberStarted Then
				strSQLwhereMemSearch = strSQLwhereMemSearch & _
				"AND (" & strDbTable & "Topic.Topic_ID " & _
					"IN (" & _
						"SELECT " & strDbTable & "Topic.Topic_ID " & _
						"FROM " & strDbTable & "Topic" & strDBNoLock & ", " & strDbTable & "Thread" & strDBNoLock & ", " & strDbTable & "Author" & strDBNoLock & " " & _
						"WHERE " & strDbTable & "Topic.Start_Thread_ID = " & strDbTable & "Thread.Thread_ID AND " & _
							strDbTable & "Thread.Author_ID = " & strDbTable & "Author.Author_ID "
			
			'Get topics that this member has posted
			Else
			
				strSQLwhereMemSearch = strSQLwhereMemSearch & _
				"AND (" & strDbTable & "Topic.Topic_ID " & _
					"IN (" & _
						"SELECT " & strDbTable & "Thread.Topic_ID " & _
						"FROM " & strDbTable & "Thread" & strDBNoLock & ", " & strDbTable & "Author" & strDBNoLock & " " & _
						"WHERE " & strDbTable & "Thread.Author_ID = " & strDbTable & "Author.Author_ID "
			End If		
			
				
			'Create the SQL for the member search, either exact match or LIKE match
			If blnExactUserMatch Then
		 		strSQLwhereMemSearch = strSQLwhereMemSearch & "AND (" & strDbTable & "Author.Username = '" & strMemberName & "') "
			Else
				strSQLwhereMemSearch = strSQLwhereMemSearch & "AND (" & strDbTable & "Author.Username LIKE '" & strMemberName & "%') "
			End If
			
			'If display hidden posts to admin, modertors, and those who posted them
			If blnModerator = false AND blnAdmin = false Then 
				strSQLwhereMemSearch = strSQLwhereMemSearch & " AND (" & strDbTable & "Thread.Hide = " & strDBFalse & " "
				If intGroupID <> 2 Then strSQLwhereMemSearch = strSQLwhereMemSearch & " OR " & strDbTable & "Thread.Author_ID = " & lngLoggedInUserID 
				strSQLwhereMemSearch = strSQLwhereMemSearch &  ") "
			End If
			strSQLwhereMemSearch = strSQLwhereMemSearch & _
				")" & _
			") "
		
			
		'Else results as shown in 'post' view so don't use the sub query (would also be faster!!)
		Else
			
		
			'Look for only posts in topics this user has started 
			'(need to use a sub query so we can get results from posts which are not the first post in the topic)
			If blnMemberStarted Then
				strSQLwhereMemSearch = strSQLwhereMemSearch & _
				"AND (" & strDbTable & "Topic.Topic_ID " & _
					"IN (" & _
						"SELECT " & strDbTable & "Topic.Topic_ID " & _
						"FROM " & strDbTable & "Topic" & strDBNoLock & ", " & strDbTable & "Thread" & strDBNoLock & ", " & strDbTable & "Author" & strDBNoLock & " " & _
						"WHERE " & strDbTable & "Topic.Start_Thread_ID = " & strDbTable & "Thread.Thread_ID AND " & _
							strDbTable & "Thread.Author_ID = " & strDbTable & "Author.Author_ID "
			
				
				'Create the SQL for the member search, either exact match or LIKE match
				'(this matches the sub query with the member we are looking for)
				If blnExactUserMatch Then
			 		strSQLwhereMemSearch = strSQLwhereMemSearch & "AND " & strDbTable & "Author.Username = '" & strMemberName & "')) "
				Else
					strSQLwhereMemSearch = strSQLwhereMemSearch & "AND " & strDbTable & "Author.Username LIKE '" & strMemberName & "%')) "
				End If
			End If
			
			
			'Create the SQL for the member search, either exact match or LIKE match
			'(this looks for any posts within the topic inwhich the user has posted)
			If blnExactUserMatch Then
		 		strSQLwhereMemSearch = strSQLwhereMemSearch & " AND (" & strDbTable & "Author.Username = '" & strMemberName & "') "
			Else
				strSQLwhereMemSearch = strSQLwhereMemSearch & " AND (" & strDbTable & "Author.Username LIKE '" & strMemberName & "%') "
			End If
		
		End If
		
	End If
	
	
	
	
	'******************************************
	'***   	SQL for date search		***
	'******************************************
	
	'If a date is selected build the SQL string for the date
	If intShowTopicsFrom <> 0 Then
		
		'Start the SQL for the date
		If strResultType = "topics" Then
			strSQLwhereDate = "AND (LastThread.Message_date"
		Else
			strSQLwhereDate = "AND (" & strDbTable & "Thread.Message_date"
		End If
		
		'Set the direction, (posts before or after date requested)
		If strDateDirection = "newer" Then
			strSQLwhereDate = strSQLwhereDate & ">"
		Else
			strSQLwhereDate = strSQLwhereDate & "<"
		End If
		
		
		'If Access use # around dates, other DB's use ' around dates
		If strDatabaseType = "Access" Then
			strSQLwhereDate = strSQLwhereDate & "#"
		Else
			strSQLwhereDate = strSQLwhereDate & "'"
		End If
		
		'Initialse the string to display when active topics are shown since
		Select Case intShowTopicsFrom
			Case 1
				strSQLwhereDate = strSQLwhereDate & internationalDateTime(dtmLastVisitDate)
			Case 2
				strSQLwhereDate = strSQLwhereDate & internationalDateTime(DateAdd("d", -1, now()))
			Case 3
				strSQLwhereDate = strSQLwhereDate & internationalDateTime(DateAdd("ww", -1, now()))
			Case 4
				strSQLwhereDate = strSQLwhereDate & internationalDateTime(DateAdd("m", -1, now()))
			Case 5
				strSQLwhereDate = strSQLwhereDate & internationalDateTime(DateAdd("m", -2, now()))
			Case 6
				strSQLwhereDate = strSQLwhereDate & internationalDateTime(DateAdd("m", -6, now()))
			Case 7
				strSQLwhereDate = strSQLwhereDate & internationalDateTime(DateAdd("yyyy", -1, now()))
		End Select
		
		'If SQL server remove dash (-) from the ISO international date to make SQL Server safe
		If strDatabaseType = "SQLServer" Then strSQLwhereDate = Replace(strSQLwhereDate, "-", "", 1, -1, 1)
		
		'If Access use # around dates, other DB's use ' around dates
		If strDatabaseType = "Access" Then
			strSQLwhereDate = strSQLwhereDate & "#"
		Else
			strSQLwhereDate = strSQLwhereDate & "'"
		End If
		
		strSQLwhereDate = strSQLwhereDate & ") "	
	End If
	

	
	'******************************************
	'***   	SQL for forum ID search		***
	'******************************************
	
	'Set the forums to look in, if not looking in all forums
	If Trim(Mid(strForumIDs, 1, 1)) <> "0" AND strForumIDs <> "" Then
		
		strSQLwhereForum = strSQLwhereForum & " AND " & strDbTable & "Forum.Forum_ID IN ("
		
		'Loop through all the selected forums
		For each iarySearchForumID in Request.Form("forumID")
			
			'If this is 2nd or more time around add OR
			If intCurrentRecord > 0 Then strSQLwhereForum = strSQLwhereForum & "," 
			
			'Add the forum ID to look in to the SQL query
			strSQLwhereForum = strSQLwhereForum & CInt(iarySearchForumID)
			
			'Add 1 to the current record position counter
			intCurrentRecord = intCurrentRecord + 1
		Next
		
		strSQLwhereForum = strSQLwhereForum & ") "
		
		'Reset record count
		intCurrentRecord = 0
	End If
	
	
	
	'******************************************
	'***   	SQL for Topic ID search		***
	'******************************************
	
	'Set the forums to look in, if not looking in all forums
	If Request.Form("qTopic") Then 
		
		strSQLwhereForum = strSQLwhereForum & " AND " & strDbTable & "Topic.Topic_ID = " & lngTopicID & " "
		
	End If
	
	
	
	'******************************************
	'***   	Main SQL Query			***
	'******************************************
	
	
	'If displaying results as topics then more data is required and a different query needs to be run
	If strResultType = "topics" Then
	
		'Initalise SQL query (quite complex but required if we only want 1 db hit to get the lot for the whole page)
		strSQL = "" & _
		"SELECT "
		If strDatabaseType = "SQLServer" OR strDatabaseType = "Access" Then
			strSQL = strSQL & " TOP " & intMaxResults & " "
		End If
		strSQL = strSQL & _
		"" & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Forum_name, " & strDbTable & "Forum.Password, " & strDbTable & "Forum.Forum_code, " & strDbTable & "Topic.Topic_ID, " & strDbTable & "Topic.Poll_ID, " & strDbTable & "Topic.Moved_ID, " & strDbTable & "Topic.Subject, " & strDbTable & "Topic.Icon, " & strDbTable & "Topic.Start_Thread_ID, " & strDbTable & "Topic.Last_Thread_ID, " & strDbTable & "Topic.No_of_replies, " & strDbTable & "Topic.No_of_views, " & strDbTable & "Topic.Locked, " & strDbTable & "Topic.Priority, " & strDbTable & "Topic.Hide, " & strDbTable & "Thread.Message_date, " & strDbTable & "Thread.Message, " & strDbTable & "Thread.Author_ID, " & strDbTable & "Author.Username, LastThread.Message_date, LastThread.Author_ID, LastAuthor.Username, " & strDbTable & "Topic.Event_date, " & strDbTable & "Topic.Event_date_end, " & strDbTable & "Topic.Rating, " & strDbTable & "Topic.Rating_Votes " & _
		"FROM " & strDbTable & "Category" & strDBNoLock & ", " & strDbTable & "Forum" & strDBNoLock & ", " & strDbTable & "Topic" & strDBNoLock & ", " & strDbTable & "Thread" & strDBNoLock & ", " & strDbTable & "Thread AS LastThread" & strDBNoLock & ", " & strDbTable & "Author" & strDBNoLock & ", " & strDbTable & "Author AS LastAuthor" & strDBNoLock & " " & _
		"WHERE " & strDbTable & "Category.Cat_ID = " & strDbTable & "Forum.Cat_ID " & _
			"AND " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Topic.Forum_ID " & _
			"AND " & strDbTable & "Thread.Author_ID = " & strDbTable & "Author.Author_ID " & _
			"AND LastThread.Author_ID = LastAuthor.Author_ID " & _
			"AND " & strDbTable & "Topic.Start_Thread_ID = " & strDbTable & "Thread.Thread_ID " & _
			"AND " & strDbTable & "Topic.Last_Thread_ID = LastThread.Thread_ID "
			
	
	
	'Else displaying results as posts so the query needs to be less complex and less data
	Else	
		
		'Initalise SQL query (quite complex but required if we only want 1 db hit to get the lot for the whole page)
		strSQL = "" & _
		"SELECT "
		If strDatabaseType = "SQLServer" OR strDatabaseType = "Access" Then
			strSQL = strSQL & " TOP " & intMaxResults & " "
		End If
		strSQL = strSQL & _
		"" & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Forum_name, " & strDbTable & "Forum.Password, " & strDbTable & "Forum.Forum_code, " & strDbTable & "Topic.Topic_ID, " & strDbTable & "Topic.Poll_ID, " & strDbTable & "Topic.Moved_ID, " & strDbTable & "Topic.Subject, " & strDbTable & "Topic.Icon, " & strDbTable & "Topic.No_of_replies, " & strDbTable & "Topic.No_of_views, " & strDbTable & "Topic.Locked, " & strDbTable & "Topic.Priority, " & strDbTable & "Topic.Hide, " & strDbTable & "Thread.Thread_ID, " & strDbTable & "Thread.Message_date, " & strDbTable & "Thread.Message, " & strDbTable & "Thread.Author_ID, " & strDbTable & "Author.Username, " & strDbTable & "Topic.Event_date, " & strDbTable & "Topic.Event_date_end, " & strDbTable & "Topic.Rating, " & strDbTable & "Topic.Rating_Votes, " & strDbTable & "Thread.IP_addr " & _
		"FROM " & strDbTable & "Category" & strDBNoLock & ", " & strDbTable & "Forum" & strDBNoLock & ", " & strDbTable & "Topic" & strDBNoLock & ", " & strDbTable & "Thread" & strDBNoLock & ", " & strDbTable & "Author" & strDBNoLock & " " & _
		"WHERE " & strDbTable & "Category.Cat_ID = " & strDbTable & "Forum.Cat_ID " & _
			"AND " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Topic.Forum_ID "
			
			'If searcing topic subjects so join the start topic ID with threads table
			If strSearhIn = "subject" Then
				strSQL = strSQL & _
				"AND " & strDbTable & "Topic.Start_Thread_ID = " & strDbTable & "Thread.Thread_ID "
			
			'Else join the topic ID with thread table
			Else
				strSQL = strSQL & _
				"AND " & strDbTable & "Topic.Topic_ID = " & strDbTable & "Thread.Topic_ID "
			End If
			
			strSQL = strSQL & _	
			"AND " & strDbTable & "Thread.Author_ID = " & strDbTable & "Author.Author_ID "
			
	End If
		
		
		
	'Put in where forum search here so we can 
	strSQL = strSQL & strSQLwhereForum
	
	
	'Check the permissions
	strSQL = strSQL & _
	"AND (" & strDbTable & "Topic.Forum_ID " & _
		"IN (" & _
			"SELECT " & strDbTable & "Permissions.Forum_ID " & _
			"FROM " & strDbTable & "Permissions " & strDBNoLock & " " & _
			"WHERE (" & strDbTable & "Permissions.Group_ID = " & intGroupID & " OR " & strDbTable & "Permissions.Author_ID = " & lngLoggedInUserID & ") " & _
				"AND " & strDbTable & "Permissions.View_Forum = " & strDBTrue & _
		")" & _
	") "
	
	
	'If this is a guest don't display hidden posts
	If intGroupID = 2 Then
		strSQL = strSQL & "AND (" & strDbTable & "Topic.Hide = " & strDBFalse & ") AND (" & strDbTable & "Thread.Hide = " & strDBFalse & ") "
	'If this isn't a moderator only display hidden posts if the user posted them
	ElseIf blnModerator = false AND blnAdmin = false Then
		strSQL = strSQL & "AND (" & strDbTable & "Topic.Hide = " & strDBFalse & " OR " & strDbTable & "Thread.Author_ID = " & lngLoggedInUserID & ") AND (" & strDbTable & "Thread.Hide=" & strDBFalse & " OR " & strDbTable & "Thread.Author_ID = " & lngLoggedInUserID & ") "
	End If
	

	
	'Place the most expenisve WHERE clauses last
	strSQL = strSQL & _
	strSQLwhereMemSearch & _
	strSQLwhereDate & _
	strSQLwhereKeywords
	
	
	'Set the sort by order
	Select Case strOrderBy
		Case "StartDate"
			'Use the Thread_ID as it should be in the same order as date but offers a massive performance boost
			strSQL = strSQL & "ORDER BY " & strDbTable & "Thread.Thread_ID ASC" 
		Case "Subject"
			strSQL = strSQL & "ORDER BY " & strDbTable & "Topic.Subject ASC"
		Case "Replies"
			strSQL = strSQL & "ORDER BY " & strDbTable & "Topic.No_of_replies DESC"
		Case "Views"
			strSQL = strSQL & "ORDER BY " & strDbTable & "Topic.No_of_views DESC"
		Case "Username"
			strSQL = strSQL & "ORDER BY " & strDbTable & "Author.Username ASC"
		Case "ForumName"
			strSQL = strSQL & "ORDER BY " & strDbTable & "Forum.Forum_name ASC"
		Case Else
			'If displaying results as topics we have created a new table using the query so results must be order by the new table
			If strResultType = "topics" Then
				'strSQL = strSQL & "ORDER BY " & strDbTable & "Topic.Last_Thread_ID DESC"
				strSQL = strSQL & "ORDER BY LastThread.Message_date DESC"
			Else
				'Use the Thread_ID as it should be in the same order as date but offers a massive performance boost
				strSQL = strSQL & "ORDER BY " & strDbTable & "Thread.Thread_ID DESC"
			End If
	End Select
	
	'mySQL limit operator
	If strDatabaseType = "mySQL" Then
		strSQL = strSQL & " LIMIT " & intMaxResults
	End If
	strSQL = strSQL & ";"



	'Response.Write(strSQL)

	
	'If the keywords and member names are not to short run the query
	If blnSearhWordsTwoShort = false Then

		'Set error trapping
		On Error Resume Next
			
		'Query the database
		rsCommon.Open strSQL, adoCon
		
		'If an error has occured then display error
		If Err.Number <> 0 Then
		 
		 	'If this is a timeout error
			If inStr(0, Err.Description, "Timeout", 1) Then
				
				blnTimeoutError = True
			Else
		
				'Display error message to page
				If strDatabaseType = "mySQL" Then	
					Call errorMsg("An error has occurred while executing SQL query on database.<br />Please check that the MySQL Server version is 4.1 or above.", "get_search_results_data", "search_process.asp")
				Else	
					Call errorMsg("An error has occurred while executing SQL query on database.", "get_search_results_data", "search_process.asp")
				End If
			End If
		
		
		'Else no error has occured
		Else
		
			'******************************************
			'***   	Save the search to app memory	***
			'******************************************
			
			'If a search result is found then save the details and recordset to memory
			If NOT rsCommon.EOF Then
			
			
				'Create an ID for the search using time in secounds for a unique value
				lngSearchID = internationalDateTime(Now())
				lngSearchID = Replace(lngSearchID, "-", "", 1, -1, 1)
				lngSearchID = Replace(lngSearchID, ":", "", 1, -1, 1)
				lngSearchID = Replace(lngSearchID, " ", "", 1, -1, 1)
			
			
				'ReDimesion the search index array
				ReDim Preserve sarySearchIndex(5, UBound(sarySearchIndex, 2) + 1)
				
				'Place the new search index data into the new array position
				sarySearchIndex(0, UBound(sarySearchIndex, 2)) = lngSearchID
				sarySearchIndex(1, UBound(sarySearchIndex, 2)) = strSeachIPAddress
				sarySearchIndex(2, UBound(sarySearchIndex, 2)) = internationalDateTime(Now())
				sarySearchIndex(3, UBound(sarySearchIndex, 2)) = lngLoggedInUserID
				sarySearchIndex(4, UBound(sarySearchIndex, 2)) = internationalDateTime(Now())
				sarySearchIndex(5, UBound(sarySearchIndex, 2)) = FormatNumber(Timer() - dblStartTime,3)
				
				
				'Place the recordset into an application array
				'Lock the application so that no other user can try and update the application level variable at the same time
				Application.Lock
				
				'Place the search record set into an array using the search ID as the application variable name
				Application(lngSearchID) = rsCommon.GetRows()
					
				'Unlock the application
				Application.UnLock
				
				'Set the results found boolean to true
				blnResultsFound = true
			End If
			
			'Close rs
			rsCommon.Close
		
		End If
		
		'Disable error trapping
		On Error goto 0
		
	End If
End If

'Clean up
Call closeDatabase()

'Update search index application array
'Lock the application so that no other user can try and update the application level variable at the same time
Application.Lock
			
'Update the application level variables
Application(strAppPrefix & "sarySearchIndex") = sarySearchIndex
			
'Unlock the application
Application.UnLock


'If the search is secsussful
If blnResultsFound Then 
	'If search results displayed in topics then redirect to topic page
	If strResultType = "topics" Then
		Response.Redirect("search_results_topics.asp?SearchID=" & lngSearchID & "&KW=" & Server.URLEncode(strSearchKeywords) & strQsSID3)
	
	'Else results displayed as posts so redirect to that page
	Else
		Response.Redirect("search_results_posts.asp?SearchID=" & lngSearchID & "&KW=" & Server.URLEncode(strSearchKeywords) & strQsSID3)
	End If
End If


'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtSearchTheForum

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title><% = strTxtSearchTheForum %></title>

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
  <td align="left"><h1><% = strTxtSearchTheForum %></h1></td>
 </tr>
</table> 
<br />
<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
 <tr>
  <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtSearchError %></strong></td>
 </tr>
 <tr>
  <td><%
 
'Write the error message
If blnSearchIPFound Then
	Response.Write("<br />" & strTxtIPSearchError)  
ElseIf blnSearhWordsTwoShort Then
	Response.Write("<br />" & strTxtSearchWordLengthError) 
ElseIf blnTimeoutError Then
	Response.Write("<br />" & strTxtSearchTimeoutPleaseNarrowSearchTryAgain) 
ElseIf blnResultsFound = false Then    
	Response.Write("<strong>" & strTxtNoSearchResultsFound & "</strong><br /><br />" & strTxtNoSearchResults) 
End If        
        %><br /><br /><a href="search_form.asp<% = strQsSID1 %>"><% = strTxtClickHereToRefineSearch %></a>
   <br />
  </td>
 </tr>
</table>
<br /><br />
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
%>
</div>
<!-- #include file="includes/footer.asp" -->