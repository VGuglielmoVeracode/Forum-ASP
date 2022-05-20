<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="admin_common.asp" -->
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





'Set the response buffer to true
Response.Buffer = True


'Dimension variables
Dim intForum		'Holds the number of fourms
Dim lngTopic		'Holds the number of topics
Dim dtmTopic		'Holds the date of the last topic
Dim lngPost		'Holds the number of posts
Dim dtmPost		'Holds the date of the last post
Dim lngPm		'Holds the number of private messages
Dim dtmPm		'Holds the date of the last private message
Dim lngPoll		'Holds the number of polls
Dim intActiveUsers	'Holds the number of active users
Dim intGroups		'Holds the number of groups
Dim lngMember		'Holds the number of members
Dim dtmMember		'Holds the date of the last members signup
Dim lngUserID		'Holds the active users ID
Dim strActUser		'Holds the active users username
Dim strForumName 	'Holds the forum name
Dim intGuestNumber	'Holds the Guest Number
Dim intActiveGuests	'Holds the number of active guests
Dim intActiveMembers	'Holds the nunber of active members
Dim strBrowserUserType	'Holds the users browser type
Dim strOS		'Holds the users OS
Dim dtmLastActive	'Holds the last active date
Dim dtmLoggedIn		'Holds the date the user logged in
Dim intArrayPass	'Loop counter
Dim strOSbrowser	'Holds the OS and browser of the user
Dim strLocation		'Holds the users location
Dim strURL		'Holds the URL to the users location
Dim blnHideActiveUser	'Holds if the user wants to be hidden
Dim strUsername		'Holds the username
Dim strDBversionInfo
Dim intDBversionNumber
Dim strLastActive


'Initilise variables
intActiveMembers = 0
intActiveGuests = 0
intActiveUsers = 0
intGuestNumber = 0
intForum = 0
lngTopic = 0
lngPost = 0
lngPm = 0
intActiveUsers = 0
intGroups = 0
lngMember = 0



'Get SQL Server version
If strDatabaseType = "SQLServer" Then
	
	'Get the sql server version from function
	strDBversionInfo = sqlServerVersion()
		
		
'Get mySQL version	
ElseIf strDatabaseType = "mySQL" Then
	
	strDBversionInfo = "mySQL " 
	
	strSQL = "SELECT VERSION() AS Version"
	rsCommon.Open strSQL, adoCon
	If NOT rsCommon.EOF Then strDBversionInfo = strDBversionInfo & rsCommon("Version")
	rsCommon.Close
End If



'Read in if active users is anbaled
blnActiveUsers = CBool(Application(strAppPrefix & "blnActiveUsers"))





'******************************************
'***	    Read in the Counts		***
'******************************************

'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT " & strDbTable & "Forum.No_of_topics, " & strDbTable & "Forum.No_of_posts FROM " & strDbTable & "Forum;"

'Query the database
rsCommon.Open strSQL, adoCon

'Get the number of topics posts and forums
Do While NOT rsCommon.EOF

 	'Count the number of forums
 	intForum = intForum + 1

 	'Count the number of topics
 	lngTopic = lngTopic + CLng(rsCommon("No_of_topics"))

 	'Count the number of posts
 	lngPost = lngPost + CLng(rsCommon("No_of_posts"))

 	'Move to the next record
 	rsCommon.MoveNext
Loop

'Clean up
rsCommon.Close



'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT Count(" & strDbTable & "Author.Author_ID) AS CountAuthor FROM " & strDbTable & "Author;"

'Query the database
rsCommon.Open strSQL, adoCon

'Read in the count
If NOT rsCommon.EOF Then lngMember = CLng(rsCommon("CountAuthor"))

'Clean up
rsCommon.Close



'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT Count(" & strDbTable & "PMMessage.PM_ID) AS CountPm FROM " & strDbTable & "PMMessage;"

'Query the database
rsCommon.Open strSQL, adoCon

'Read in the count
If NOT rsCommon.EOF Then lngPm = CLng(rsCommon("CountPm"))

'Clean up
rsCommon.Close


'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT Count(" & strDbTable & "Poll.Poll_ID) AS CountPoll FROM " & strDbTable & "Poll;"

'Query the database
rsCommon.Open strSQL, adoCon

'Read in the count
If NOT rsCommon.EOF Then lngPoll = CLng(rsCommon("CountPoll"))

'Clean up
rsCommon.Close



'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT Count(" & strDbTable & "Group.Group_ID) AS CountGroup FROM " & strDbTable & "Group;"

'Query the database
rsCommon.Open strSQL, adoCon

'Read in the count
If NOT rsCommon.EOF Then intGroups = CLng(rsCommon("CountGroup"))

'Clean up
rsCommon.Close



'******************************************
'***	    	Read in Dates		***
'******************************************

'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT" & strDBTop1 & " " & strDbTable & "Thread.Message_date " & _
"FROM " & strDbTable & "Thread, " & strDbTable & "Topic " & _
"WHERE " & strDbTable & "Thread.Thread_ID = " & strDbTable & "Topic.Start_Thread_ID " & _
"ORDER BY " & strDbTable & "Thread.Message_date DESC" & strDBLimit1 & ";"

'Query the database
rsCommon.Open strSQL, adoCon

'Read in the count
If NOT rsCommon.EOF Then dtmTopic = CDate(rsCommon("Message_date"))

'Clean up
rsCommon.Close



'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT" & strDBTop1 & " " & strDbTable & "Thread.Message_date FROM " & strDbTable & "Thread ORDER BY " & strDbTable & "Thread.Message_date DESC" & strDBLimit1 & ";"

'Query the database
rsCommon.Open strSQL, adoCon

'Read in the count
If NOT rsCommon.EOF Then dtmPost = CDate(rsCommon("Message_date"))

'Clean up
rsCommon.Close



'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT" & strDBTop1 & " " & strDbTable & "Author.Join_date FROM " & strDbTable & "Author ORDER BY " & strDbTable & "Author.Join_date DESC" & strDBLimit1 & ";"

'Query the database
rsCommon.Open strSQL, adoCon

'Read in the count
If NOT rsCommon.EOF Then dtmMember = CDate(rsCommon("Join_date"))

'Clean up
rsCommon.Close




'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT" & strDBTop1 & " " & strDbTable & "PMMessage.PM_Message_date FROM " & strDbTable & "PMMessage ORDER BY " & strDbTable & "PMMessage.PM_Message_date DESC" & strDBLimit1 & ";"

'Query the database
rsCommon.Open strSQL, adoCon

'Read in the count
If NOT rsCommon.EOF Then dtmPm = CDate(rsCommon("PM_Message_date"))

'Clean up
rsCommon.Close


%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Forum Statistics</title>
<meta name="generator" content="Web Wiz Forums" />
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
  <h1>Forum Statistics</h1><br />
    <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br /></p>
    <br />
    <br />
</div>
<table align="center" cellpadding="4" cellspacing="1" class="tableBorder">
  <tr align="left">
    <td colspan="4" class="tableLedger">Forum Statistics</td>
  </tr>
  <tr>
   <td  height="2" align="left" class="tableRow">Forum Version </td>
   <td valign="top" class="tableRow"><% = strVersion %></td>
   <td valign="top" class="tableRow">Updates</td>
   <td height="2" valign="top" class="tableRow"><a href="admin_update_check.asp">Check for updates</a></td>
 </tr>
  <tr>
   <td  height="2" align="left" class="tableRow" valign="top">Database</td>
   <td valign="top" class="tableRow" colspan="3"><% If strDBversionInfo <> "" Then Response.Write(strDBversionInfo) Else Response.Write(strDatabaseType) %></td>
  </tr>
  <tr>
    <td width="31%" align="left" class="tableRow">Number of Forums<span class="smText"></span></td>
    <td valign="top"  colspan="3" class="tableRow"><% = intForum %>    </td>
  </tr>
  <tr>
    <td width="31%" align="left" class="tableRow">Number of Topics<span class="smText"></span></td>
    <td width="15%" valign="top" class="tableRow"><% = lngTopic %>    </td>
   <td width="22%" valign="top" class="tableRow">Last New Topic</td>
   <td width="32%" height="12" valign="top" class="tableRow"><% = DateFormat(dtmTopic) & ", " &  TimeFormat(dtmTopic) %>    </td>
  </tr>
  <tr>
    <td width="31%" align="left" class="tableRow">Number of Posts<span class="smText"></span></td>
    <td width="15%" valign="top" class="tableRow"><% = lngPost %>    </td>
   <td width="22%" valign="top" class="tableRow">Last New Post</td>
   <td width="32%" height="12" valign="top" class="tableRow"><% = DateFormat(dtmPost) & ", " &  TimeFormat(dtmPost) %>    </td>
  </tr>
  <tr>
    <td  height="2" align="left" class="tableRow">Number of Members</td>
    <td valign="top" class="tableRow"><% = lngMember %>    </td>
    <td valign="top" class="tableRow">Last New Member</td>
    <td height="2" valign="top" class="tableRow"><% = DateFormat(dtmMember) & ", " &  TimeFormat(dtmMember) %>    </td>
  </tr>
  <tr>
    <td  height="2" align="left" class="tableRow">Number of Private Messages</td>
    <td valign="top" class="tableRow"><% = lngPm %>    </td>
    <td valign="top" class="tableRow">Last Private Message</td>
    <td height="2" valign="top" class="tableRow"><% = DateFormat(dtmPm) & ", " &  TimeFormat(dtmPm) %>    </td>
  </tr>
  <tr>
    <td  height="2" align="left" class="tableRow">Number of Polls</td>
    <td valign="top"  colspan="3" class="tableRow"><% = lngPoll %>    </td>
  </tr>
  <tr>
   <td  height="2" align="left" class="tableRow">Number of User Groups</td>
   <td valign="top"  colspan="3" class="tableRow"><% = intGroups %>   </td>
  </tr><%

If blnActiveUsers Then
	
%>
  <tr>
   <td  height="2" align="left" class="tableRow">Most Active Users</td>
   <td valign="top" class="tableRow"><% = lngMostEverActiveUsers %></td>
   <td valign="top" class="tableRow">Most Active Date</td>
   <td height="2" valign="top" class="tableRow"><% = DateFormat(dtmMostEverActiveDate) & ", " &  TimeFormat(dtmMostEverActiveDate) %></td>
  </tr><%
End If

%>
</table>
<br />
<br />
<br /><%

If blnActiveUsers Then

	'Initialise  the array from the application veriable
	If IsArray(Application(strAppPrefix & "saryAppActiveUsersTable")) Then 
			
		'Place the application level active users array into a temporary dynaimic array
		saryActiveUsers = Application(strAppPrefix & "saryAppActiveUsersTable")
		
	'Else Initialise the an empty array
	Else
		ReDim saryActiveUsers(8,0)
	End If
	
	
	'Sort the active users array
	Call SortActiveUsersList(saryActiveUsers)
	
	
	'Get the number of active users
	'Get the active users online
	For intArrayPass = 1 To UBound(saryActiveUsers, 2)
		
		'If this is a guest user then increment the number of active guests veriable
		If saryActiveUsers(1, intArrayPass) = 2 Then 	
				
			intActiveGuests = intActiveGuests + 1
		End If
			
	Next 
	
	'Calculate the number of members online and total people online
	intActiveUsers = UBound(saryActiveUsers, 2)
	intActiveMembers = intActiveUsers - intActiveGuests

%>    
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left">
    <div style="float:left;"><% = strTxtThereAreCurrently & " " & intActiveUsers & " " & strTxtActiveUsers & " " & strTxtOnLine & ", "  & intActiveGuests & " " & strTxtGuests & " and " & intActiveMembers & " " & strTxtMembers %></div>
    <div style="float:right;"><a href="admin_statistics.asp<% = strQsSID1 %>"><img src="<% = strImagePath %>refresh.png" alt="<% = strTxtRefreshPage %>" title="<% = strTxtRefreshPage %>" /></a>&nbsp;</div>
  </td>
 </tr>
<table>
<table cellspacing="1" cellpadding="3" class="tableBorder" align="center">
 <tr class="tableLedger">
  <td width="14%"><% = strTxtUsername %></td>
  <td width="10%">IP</td>
  <td width="14%"><% = strTxtLastActive %></td>
  <td width="10%"><% = strTxtActive %></td>
  <td width="22%" align="left"><% = strTxtOS & "/" & strTxtBrowser %></td>
  <td width="33%" align="left"><% = strTxtLocation %></td>
 </tr><%  


        		
	'display the active users
	For intArrayPass = 1 To UBound(saryActiveUsers, 2)
		
		'Array dimension lookup table
		' 0 = IP
		' 1 = Author ID
		' 2 = Username
		' 3 = Login Time
		' 4 = Last Active Time
		' 5 = OS/Browser
		' 6 = Location Page Name
		' 7 = URL
		' 8 = Hids user details
	
		'Read in the details from the rs
		lngUserID = saryActiveUsers(1, intArrayPass)
		strUsername = saryActiveUsers(2, intArrayPass)
		dtmLoggedIn = saryActiveUsers(3, intArrayPass)
		dtmLastActive = saryActiveUsers(4, intArrayPass)
		strOSbrowser = saryActiveUsers(5, intArrayPass)
		strLocation = saryActiveUsers(6, intArrayPass)
		strURL = saryActiveUsers(7, intArrayPass)
		blnHideActiveUser = CBool(saryActiveUsers(8, intArrayPass))
		
		'Setup the last active date
		strLastActive = DateFormat(dtmLoggedIn)
		strLastActive = Replace(strLastActive, "<strong>", "")
		strLastActive = Replace(strLastActive, "</strong>", "")
		
				
		'Write the HTML of the Topic descriptions as hyperlinks to the Topic details and message
%>
 <tr class="tableRow"> 
  <td><% 
          
	         'If the user is a Guest then display them as a Guest
	         If lngUserID = 2 Then
	         
	         	'Add 1 to the Guest number
	         	intGuestNumber = intGuestNumber + 1
	         	
	         	'Display the User as Guest
	         	Response.Write(strTxtGuest & " "& intGuestNumber)
	         
	         'Else display the users name
	         Else 
	         
	         %><a href="member_profile.asp?PF=<% = lngUserID %>" rel="nofollow"><% = strUsername %></a><% 
	        
	        End If
        
        %></td>
  <td nowrap><% Response.Write(saryActiveUsers(0, intArrayPass))  %></td>
  <td nowrap><% Response.Write(strLastActive & " " & strTxtAt & "&nbsp;" & TimeFormat(dtmLoggedIn))  %></td>
  <td><% = DateDiff("n", dtmLoggedIn, dtmLastActive) %>&nbsp;<% = strTxtMinutes %></td>
  <td nowrap><% = strOSbrowser %></td>
  <td nowrap><% = strLocation %><% If strLocation <> "" AND strURL <> "" Then Response.Write("<br />") %><% = strURL %></td>
 </tr><%
		
	   		
	Next
	


%>
</table>
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td align="center" class="text"><br />
        This data is based on users active over the past twenty minutes</td>
    </tr>
  </table>
  <br />
  <br />
  <%
End If

'Clean up
Call closeDatabase()

%>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
