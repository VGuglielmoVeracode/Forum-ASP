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




Server.ScriptTimeout = 2000000 'secounds


'Set the response buffer to true as we maybe redirecting
Response.Buffer = True


'If in demo mode redirect
If blnDemoMode Then
	Call closeDatabase()
	Response.Redirect("admin_web_wiz_forums_premium.asp" & strQsSID1)
End If


'Dimension variables
Dim intNoOfDays			'Holds the number of days to delete posts from
Dim lngNumberOfTopics		'Holds the number of topics that are deleted
Dim lngPollID			'Holds the poll ID if there is one to delete
Dim rsThread			'Holds the threads recordset
Dim intPriority			'Holds the topic priority to delete
Dim saryFileUploads		'Holds the files to be deleted
Dim intLoop			'Loop counter
Dim objFSO			'Holds the FSO object
Dim dtmSelectedDate

'Initilise variables
lngNumberOfTopics = 0



'Check the form ID to prevent XCSRF
Call checkFormID(Request.Form("formID"))


'get teh number of days to delte from
intNoOfDays = IntC(Request.Form("days"))
intForumID = IntC(Request.Form("FID"))
intPriority = IntC(Request.Form("priority"))



dtmSelectedDate = internationalDateTime(DateAdd("d", -intNoOfDays, now()))

'If SQL server remove dash (-) from the ISO international date to make SQL Server safe
If strDatabaseType = "SQLServer" Then dtmSelectedDate = Replace(dtmSelectedDate, "-", "", 1, -1, 1)

'If acess used # around dates
If strDatabaseType = "Access" Then
	dtmSelectedDate = "#" & dtmSelectedDate & "#"
Else
	dtmSelectedDate = "'" & dtmSelectedDate & "'"
End If


'Get all the Topics from the database to be deleted

'Initalise the strSQL variable with an SQL statement to get the topic from the database
strSQL = "SELECT " & strDbTable & "Topic.Topic_ID, " & strDbTable & "Topic.Poll_ID " & _
"FROM " & strDbTable & "Topic, " & strDbTable & "Thread "
If intForumID = 0 Then
	strSQL = strSQL & "WHERE (" & strDbTable & "Topic.Last_Thread_ID = " & strDbTable & "Thread.Thread_ID) AND " & strDbTable & "Thread.Message_date < " & dtmSelectedDate & " "
Else
	strSQL = strSQL & "WHERE (" & strDbTable & "Topic.Last_Thread_ID = " & strDbTable & "Thread.Thread_ID) AND  (" & strDbTable & "Thread.Message_date < " & dtmSelectedDate & ") AND (" & strDbTable & "Topic.Forum_ID=" & intForumID & ") "
End If
If intPriority <> 4 Then strSQL = strSQL & " AND (" & strDbTable & "Topic.Priority=" & intPriority & ")"
strSQL = strSQL & ";"

'Query the database
rsCommon.Open strSQL, adoCon


'Create a record set object to the Threads held in the database
Set rsThread = Server.CreateObject("ADODB.Recordset")


'Loop through all the topics to get all the thread in the topics to be deleted
Do While NOT rsCommon.EOF

	'Update the number of topics deletd
	lngNumberOfTopics = lngNumberOfTopics + 1
	
	'Get the Poll ID
	lngPollID = CLng(rsCommon("Poll_ID"))
	
	
	
	'See if there are any guest posters and delete thier names form the guest name table
	
	'Initalise the strSQL variable with an SQL statement to get the topic from the database
	strSQL = "SELECT " & strDbTable & "Thread.Thread_ID " & _
	"FROM " & strDbTable & "Thread " & _
	"WHERE " & strDbTable & "Thread.Topic_ID = " & rsCommon("Topic_ID") & ";"
	
	'Query the database
	rsThread.Open strSQL, adoCon
	
	'Loop through and delete al names in the guest name table
	Do While NOT rsThread.EOF
		
		'First we need to delete any entry in the GuestName table incase this was a guest poster posting the message
		strSQL = "DELETE FROM " & strDbTable & "GuestName WHERE " & strDbTable & "GuestName.Thread_ID=" & CLng(rsThread("Thread_ID")) & ";"
		
		'Excute SQL
		adoCon.Execute(strSQL)
		
		'Move next
		rsThread.MoveNext
	Loop
	
	'Close the rs
	rsThread.Close
	
	

	'Delete the thread
	strSQL = "DELETE FROM " & strDbTable & "Thread WHERE " & strDbTable & "Thread.Topic_ID =" & rsCommon("Topic_ID") & ";"

	'Delete the threads
	adoCon.Execute(strSQL)
	
	

	'If there is a poll delete that as well
	If lngPollID > 0 Then

		'Delete the poll choice
		strSQL = "DELETE FROM " & strDbTable & "PollChoice WHERE " & strDbTable & "PollChoice.Poll_ID =" & lngPollID & ";"

		'Delete the threads
		adoCon.Execute(strSQL)

		'Delete the poll choice
		strSQL = "DELETE FROM " & strDbTable & "Poll WHERE " & strDbTable & "Poll.Poll_ID =" & lngPollID & ";"

		'Delete the threads
		adoCon.Execute(strSQL)
	End If
	
	'delete any rating for this topic
	strSQL = "DELETE FROM " & strDbTable & "TopicRatingVote " & strRowLock & " " & _
	"WHERE " & strDbTable & "TopicRatingVote.Topic_ID = " & rsCommon("Topic_ID") & ";"
		
	'Excute SQL
	adoCon.Execute(strSQL)
	
	
	'Delete the topic
	strSQL = "DELETE FROM " & strDbTable & "Topic " & _
	"WHERE " & strDbTable & "Topic.Topic_ID =" & rsCommon("Topic_ID") & ";"

	'Delete the threads
	adoCon.Execute(strSQL)

	'Move to the next record
	rsCommon.MoveNext
Loop




'Update post count
updateForumStats(intForumID)


'Reset Server Objects
Set rsThread = Nothing
rsCommon.Close
Call closeDatabase()

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title>Batch Delete Forum Topics</title>

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
  
<!-- #include file="includes/admin_header_inc.asp" -->
<div align="center">
 <h1>Batch Delete Forum Topics </h1><br />
  <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
  <br>
  <br>
  <br>
  <br>
  <br />
  <span>
  <% = lngNumberOfTopics %>
  Topics have been Deleted.</span><br />
  <br>
  <br>
  <br>
  <br>
  <a href="admin_resync_forum_post_count.asp<% = strQsSID1 %>">Click here to re-sync Post and Topic Counts for the Forums</a> </p>
</div>
<!-- #include file="includes/admin_footer_inc.asp" -->
