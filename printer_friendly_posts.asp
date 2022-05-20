<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="functions/functions_date_time_format.asp" -->
<!--#include file="includes/emoticons_inc.asp" -->
<!--#include file="functions/functions_format_post.asp" -->
<!--#include file="functions/functions_edit_post.asp" -->
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
Dim strForumDescription		'Holds the description of the forum
Dim strCatName			'Holds the cat name
Dim lngTopicID			'Holds the topic number
Dim lngMessageID		'Holds the message ID number
Dim strSubject			'Holds the topic subject
Dim strUsername 		'Holds the Username of the thread
Dim lngUserID			'Holds the ID number of the user
Dim strAuthorSignature		'Holds the authors signature
Dim dtmTopicDate		'Holds the date the thread was made
Dim strMessage			'Holds the message body of the thread
Dim blnForumLocked		'Set to true if the forum is locked
Dim lngPollID			'Holds the poll ID
Dim intCurrentRecord		'Holds the current records for the posts
Dim strGuestUsername		'Holds the Guest Username if it is a guest posting
Dim lngTotalRecords		'Holds the total number of therads in this topic


'Initialise variables
lngTopicID = 0	
intForumID = 0


'Read in the Forum ID to display the Topics for
If isNumeric(Request.QueryString("TID")) Then lngTopicID = LngC(Request.QueryString("TID")) Else lngTopicID = 0


'If there no Topic ID then redirect the user to the main forum page
If lngTopicID = 0 Then

	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("default.asp" & strQsSID1)
End If


'Get the posts from the database

strSQL = "" & _
"SELECT" & strDBTop1 & " " & strDbTable & "Category.Cat_name, " & strDbTable & "Forum.Forum_ID AS ForumID, " & strDbTable & "Forum.Forum_name, " & strDbTable & "Forum.Forum_description, " & strDbTable & "Forum.Password, " & strDbTable & "Forum.Forum_code, " & strDbTable & "Topic.Subject, " & strDbTable & "Permissions.View_Forum  " & _
"FROM " & strDbTable & "Category" & strDBNoLock & ", " & strDbTable & "Forum" & strDBNoLock & ", " & strDbTable & "Forum AS " & strDbTable & "Forum2" & strDBNoLock & ", " & strDbTable & "Permissions" & strDBNoLock & ", " & strDbTable & "Topic" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Category.Cat_ID = " & strDbTable & "Forum.Cat_ID " & _
	"AND " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Permissions.Forum_ID " & _
	"AND " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Topic.Forum_ID " & _
	"AND " & strDbTable & "Topic.Topic_ID = " & lngTopicID & " " & _
	"AND (" & strDbTable & "Permissions.Author_ID = " & lngLoggedInUserID & " OR " & strDbTable & "Permissions.Group_ID = " & intGroupID & ") " & _
"ORDER BY " & strDbTable & "Forum.Sub_ID, " & strDbTable & "Forum.Forum_Order, " & strDbTable & "Permissions.Author_ID DESC" & strDBLimit1 & ";"


'Query the database
rsCommon.Open strSQL, adoCon


'If there is no record returended then set a message to say that
If rsCommon.EOF Then
	
	'If there are no thread's to display then display the appropriate error message
	strSubject = strTxtNoThreads



'Else get the details of the forum, permissions, and topic details
Else
	
	'Read in forum details from the database
	intForumID = Cint(rsCommon("ForumID"))
	strCatName = rsCommon("Cat_name")
	strForumName = rsCommon("Forum_name")
	strForumDescription = rsCommon("Forum_description")
	
	'Read in the forum permissions
	blnRead = CBool(rsCommon("View_Forum"))
	
	'Read in the topic details
	strSubject = rsCommon("Subject")
	
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
	If rsCommon("Password") <> "" AND (getCookie("fID", "Forum" & intForumID) <> rsCommon("Forum_code") AND getSessionItem("FP" & intForumID) <> "1") Then

		'Reset Server Objects
		rsCommon.Close
		Call closeDatabase()

		'Redirect to a page asking for the user to enter the forum password
		Response.Redirect("forum_password_form.asp?RP=PT&FID=" & intForumID & "&TID=" & lngTopicID & strQsSID3)
	End If
End If

'clean up
rsCommon.Close



'Intilise SQL query to get all the posts
'Use a LEFT JOIN for the Guest name as there may not be a Guest name and so we want to include null values 	
strSQL = "" & _
"SELECT "
If strDatabaseType = "SQLServer" OR strDatabaseType = "Access" Then
	strSQL = strSQL & " TOP 100 "
End If
strSQL = strSQL & _
" " &  strDbTable & "Thread.Thread_ID, " & strDbTable & "Thread.Message, " & strDbTable & "Thread.Message_date, " & strDbTable & "Thread.Show_signature, " & strDbTable & "Thread.IP_addr, " & strDbTable & "Thread.Hide, " & strDbTable & "Author.Author_ID, " & strDbTable & "Author.Username, " & strDbTable & "Author.Homepage, " & strDbTable & "Author.Location, " & strDbTable & "Author.No_of_posts, " & strDbTable & "Author.Join_date, " & strDbTable & "Author.Signature, " & strDbTable & "Author.Active, " & strDbTable & "Author.Avatar, " & strDbTable & "Author.Avatar_title, " & strDbTable & "Group.Name, " & strDbTable & "Group.Stars, " & strDbTable & "Group.Custom_stars, " & strDbTable & "GuestName.Name " & _
"FROM (" & strDbTable & "Group INNER JOIN (" & strDbTable & "Author INNER JOIN " & strDbTable & "Thread ON " & strDbTable & "Author.Author_ID = " & strDbTable & "Thread.Author_ID) ON " & strDbTable & "Group.Group_ID = " & strDbTable & "Author.Group_ID) LEFT JOIN " & strDbTable & "GuestName ON " & strDbTable & "Thread.Thread_ID = " & strDbTable & "GuestName.Thread_ID " & _
"WHERE " & strDbTable & "Thread.Topic_ID = " & lngTopicID & " "	& _
	"AND " & strDbTable & "Thread.Hide = " & strDBFalse & " " & _
"ORDER BY " & strDbTable & "Thread.Message_date ASC"
'mySQL limit operator
If strDatabaseType = "mySQL" Then
	strSQL = strSQL & " LIMIT 100"
End If
strSQL = strSQL & ";"
		

'Query the database
rsCommon.Open strSQL, adoCon


'If there is a topic in the database then get the post data
If NOT rsCommon.EOF Then
	
	'Read in the topivc recordset into an array
	sarryPosts = rsCommon.GetRows()
	
	'SQL Query Array Look Up table
	'0 = tblThread.Thread_ID, 
	'1 = tblThread.Message, 
	'2 = tblThread.Message_date, 
	'3 = tblThread.Show_signature, 
	'4 = tblThread.IP_addr, 
	'5 = tblThread.Hide, 
	'6 = tblAuthor.Author_ID,
	'7 = tblAuthor.Username, 
	'8 = tblAuthor.Homepage, 
	'9 = tblAuthor.Location, 
	'10 = tblAuthor.No_of_posts, 
	'11 = tblAuthor.Join_date, 
	'12 = tblAuthor.Signature, 
	'13 = tblAuthor.Active, 
	'14 = tblAuthor.Avatar, 
	'15 = tblAuthor.Avatar_title, 
	'16 = tblGroup.Name, 
	'17 = tblGroup.Stars, 
	'18 = tblGroup.Custom_stars
	'19 = tblGuestName.Name
	
	'Count the number of records
	lngTotalRecords = Ubound(sarryPosts,2) + 1
End If


'Clean up
rsCommon.Close


'If active users is enabled update the active users application array
If blnActiveUsers Then
	'Call active users function
	saryActiveUsers = activeUsers(strTxtViewingTopic, strSubject, "forum_posts.asp?TID=" & lngTopicID, intForumID)
End If
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title><% = strForumName & " - " & strSubject %></title>
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
<meta name="robots" content="noindex, nofollow" />
<link href="<% = strCSSfile %>printer_style.css" rel="stylesheet" type="text/css" />
</head>
<body>
<table width="98%" border="0" cellspacing="0" cellpadding="1" align="center">
  <tr>
    <td align="center"><a href="javascript:onclick=window.print()"><% = strTxtPrintPage %></a> | <a href="JavaScript:onclick=window.close()"><% = strTxtCloseWindow %></a></td>
  </tr>
</table>
<table width="98%" border="0" cellspacing="0" cellpadding="1" align="center">
  <tr>
    <td class="smText"> <br />
      <%
	      

'If there are no threads returned by the qury then display an error message
If lngTotalRecords = 0 Then
%>
      <table width="100%" border="0" cellspacing="0" cellpadding="1">
        <tr>
          <td align="center" height="66" class="text"><strong><% = strTxtNoThreads %></strong></td>
        </tr>
      </table><%
'Display the threads
Else

	'Read in threads details for the topic from the database
	lngMessageID = CLng(sarryPosts(0,intCurrentRecord))
	strMessage = sarryPosts(1,intCurrentRecord)
	dtmTopicDate = CDate(sarryPosts(2,intCurrentRecord))
	lngUserID = CLng(sarryPosts(6,intCurrentRecord))
	strUsername = sarryPosts(7,intCurrentRecord)
	strAuthorSignature = sarryPosts(12,intCurrentRecord)
	strGuestUsername = sarryPosts(19,intCurrentRecord)
	

	'If the message has been edited remove who edited the post
	If InStr(1, strMessage, "<edited>", 1) Then strMessage = removeEditorAuthor(strMessage)

	'Convert message to text
	strMessage = ConvertToText(strMessage)

	'If the post contains a quote or code block then format it
	If InStr(1, strMessage, "[QUOTE=", 1) > 0 AND InStr(1, strMessage, "[/QUOTE]", 1) > 0 Then strMessage = formatUserQuote(strMessage)
	If InStr(1, strMessage, "[QUOTE]", 1) > 0 AND InStr(1, strMessage, "[/QUOTE]", 1) > 0 Then strMessage = formatQuote(strMessage)
	If InStr(1, strMessage, "[CODE]", 1) > 0 AND InStr(1, strMessage, "[/CODE]", 1) > 0 Then strMessage = formatCode(strMessage)
	If InStr(1, strMessage, "[HIDE]", 1) > 0 AND InStr(1, strMessage, "[/HIDE]", 1) > 0 Then strMessage = formatHide(strMessage)
	If InStr(1, strMessage, "[SPOILER]", 1) > 0 AND InStr(1, strMessage, "[/SPOILER]", 1) > 0 Then strMessage = formatSpoiler(strMessage, True)


	'If the post contains a flash link then format it
	If blnFlashFiles Then
		If InStr(1, strMessage, "[FLASH", 1) > 0 AND InStr(1, strMessage, "[/FLASH]", 1) > 0 Then strMessage = formatFlash(strMessage)
		If InStr(1, strAuthorSignature, "[FLASH", 1) > 0 AND InStr(1, strAuthorSignature, "[/FLASH]", 1) > 0 Then strAuthorSignature = formatFlash(strAuthorSignature)
	End If
	
	'If YouTube
	If blnYouTube Then
		'YouTube
		If InStr(1, strMessage, "[TUBE]", 1) > 0 AND InStr(1, strMessage, "[/TUBE]", 1) > 0 Then strMessage = formatYouTube(strMessage)
		If InStr(1, strAuthorSignature, "[TUBE]", 1) > 0 AND InStr(1, strAuthorSignature, "[/TUBE]", 1) > 0 Then strMessage = formatYouTube(strAuthorSignature)
	End If

	'If the user wants there signature shown then attach it to the message
	If CBool(sarryPosts(3,intCurrentRecord)) AND NOT strAuthorSignature = "" Then 
		strAuthorSignature = ConvertToText(strAuthorSignature)
		strMessage = strMessage & "<!-- Signature --><br /><br />-------------<br />" & strAuthorSignature
	End If

    %>
      <strong style="font-size: 16px;"><% = strSubject %></strong> <br />
      <br />
      <strong><% = strTxtPrintedFrom %>: </strong><% = strWebsiteName %>
       <br /><strong><% = strTxtCategory %>: </strong> <% = strCatName %>
      <br /><strong><% = strTxtForumName %>: </strong> <% = strForumName %>
      <br /><strong><% = strTxtForumDiscription %>: </strong> <% = strForumDescription %>
      <br /><strong><% = strTxtURL %>: </strong><a href="<% = strForumPath %>forum_posts.asp?TID=<% = lngTopicID %>"><% = strForumPath %>forum_posts.asp?TID=<% = lngTopicID %></a>
      <br /><strong><% = strTxtPrintedDate %>: </strong><% = stdDateFormat(Now(), True) & " " & strTxtAt & " " & TimeFormat(Now()) %>
      <% If blnLCode = True Then %><br /><strong><% = strTxtSoftwareVersion %>:</strong> Web Wiz Forums <% = strVersion %> - http://www.webwizforums.com<% End If %>
      <br /><br /><br />
      <span class="text"><strong><% = strTxtTopic %>:</strong> <% = strSubject %></span>
      <hr style="border-top-width: 1px" />
      <strong><% = strTxtPostedBy %>:</strong> <% = strUsername %>
      <br /><strong><% = strTxtSubjectFolder %>:</strong> <% = strSubject %>
      <br /><strong><% = strTxtDatePosted %>:</strong> <% = stdDateFormat(dtmTopicDate, True) %>&nbsp;<% = strTxtAt %>&nbsp;<% = TimeFormat(dtmTopicDate) %>
      <hr />
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td class="text">
<%  Response.Write("	<!-- Message body -->" & vbCrLf & strMessage & vbCrLf &  "<!-- Message body ''"""" -->") %>
         </td>
        </tr>
       </table>
      <br /><%
      	
	'If there are replies then display that there are
      	If lngTotalRecords > 1 Then Response.Write "<hr /><br /><strong>" & strTxtReplies & ": </strong>"
      		
      	'Move to the next record
	intCurrentRecord = intCurrentRecord + 1
      	%>
      <hr style="border-top-width: 1px" /><%


      	'Do....While Loop to loop through the recorset to display the topic posts
	Do While intCurrentRecord < lngTotalRecords
	
		'If there are no post records left to display then exit loop
		If intCurrentRecord >= lngTotalRecords Then Exit Do

		'Read in threads details for the topic from the database
		lngMessageID = CLng(sarryPosts(0,intCurrentRecord))
		strMessage = sarryPosts(1,intCurrentRecord)
		dtmTopicDate = CDate(sarryPosts(2,intCurrentRecord))
		lngUserID = CLng(sarryPosts(6,intCurrentRecord))
		strUsername = sarryPosts(7,intCurrentRecord)
		strAuthorSignature = sarryPosts(12,intCurrentRecord)
		strGuestUsername = sarryPosts(19,intCurrentRecord)
		
		

		'If the message has been edited remove who edited the post
		If InStr(1, strMessage, "<edited>", 1) Then strMessage = removeEditorAuthor(strMessage)

		'Convert message to text
		strMessage = ConvertToText(strMessage)
		

		'If the post contains a quote or code block then format it
		If InStr(1, strMessage, "[QUOTE=", 1) > 0 AND InStr(1, strMessage, "[/QUOTE]", 1) > 0 Then strMessage = formatUserQuote(strMessage)
		If InStr(1, strMessage, "[QUOTE]", 1) > 0 AND InStr(1, strMessage, "[/QUOTE]", 1) > 0 Then strMessage = formatQuote(strMessage)
		If InStr(1, strMessage, "[CODE]", 1) > 0 AND InStr(1, strMessage, "[/CODE]", 1) > 0 Then strMessage = formatCode(strMessage)
		If InStr(1, strMessage, "[SPOILER]", 1) > 0 AND InStr(1, strMessage, "[/SPOILER]", 1) > 0 Then strMessage = formatSpoiler(strMessage, True)


		'If the post contains a flash link then format it
		If blnFlashFiles Then
			If InStr(1, strMessage, "[FLASH", 1) > 0 AND InStr(1, strMessage, "[/FLASH]", 1) > 0 Then strMessage = formatFlash(strMessage)
			If InStr(1, strAuthorSignature, "[FLASH", 1) > 0 AND InStr(1, strAuthorSignature, "[/FLASH]", 1) > 0 Then strAuthorSignature = formatFlash(strAuthorSignature)
		End If
		
		'If YouTube
		If blnYouTube Then
			'YouTube
			If InStr(1, strMessage, "[TUBE]", 1) > 0 AND InStr(1, strMessage, "[/TUBE]", 1) > 0 Then strMessage = formatYouTube(strMessage)
			If InStr(1, strAuthorSignature, "[TUBE]", 1) > 0 AND InStr(1, strAuthorSignature, "[/TUBE]", 1) > 0 Then strMessage = formatYouTube(strAuthorSignature)
		End If
		
		'If the user wants there signature shown then attach it to the message
		If CBool(sarryPosts(3,intCurrentRecord)) AND NOT strAuthorSignature = "" Then
			strAuthorSignature = ConvertToText(strAuthorSignature)
			strMessage = strMessage & "<!-- Signature --><br /><br />-------------<br />" & strAuthorSignature
		End If

	      %>
      <strong><% = strTxtPostedBy %>:</strong> <% = strUsername %>
      <br />
      <strong><% = strTxtDatePosted %>:</strong> <% = stdDateFormat(dtmTopicDate, True) %>&nbsp;<% = strTxtAt %>&nbsp;<% = TimeFormat(dtmTopicDate) %>
      <hr />
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td class="text">
<%  Response.Write("	<!-- Message body -->" & vbCrLf & strMessage & vbCrLf &  "<!-- Message body ''"""" -->") %>
         </td>
        </tr>
       </table>
       <br />
      <hr style="border-top-width: 1px" />
      <%
	      	'Move to the next record
		intCurrentRecord = intCurrentRecord + 1
	Loop
%>
    </td>
  </tr>
</table>
<br />
<%
End If


'Clean up
Call closeDatabase()

%>
<table width="98%" border="0" cellspacing="0" cellpadding="1" align="center">
  <tr>
    <td align="center"><a href="javascript:onclick=window.print()"><% = strTxtPrintPage %></a> | <a href="JavaScript:onclick=window.close()"><% = strTxtCloseWindow %></a>
    <br /><br /><%
     
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnLCode = True Then
	Response.Write("<span class=""text"" style=""font-size:10px"">Forum Software by Web Wiz Forums&reg; version " & strVersion& " - http://www.webwizforums.com</span>")
	Response.Write("<br /><span class=""text"" style=""font-size:10px"">Copyright &copy;2001-2019 Web Wiz Ltd. - https://www.webwiz.net</span>")
End If 
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******

%>
</td>
  </tr>
</table>
</body>
</html>