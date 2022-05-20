<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="language_files/admin_language_file_inc.asp" -->
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


'Dimension variables
Dim strTopicSubject		'Holds the name of the topic
Dim lngTopicID			'Holds the ID number of the category
Dim lngMoveToForumID		'Holds the forum id to jump to
Dim lngPostID			'Holds the post ID
Dim intCurrentRecord		'Holds the recordset array position
Dim sarryForumSubjects

'If the user is user is using a banned IP redirect to an error page
If bannedIP() Then
	
	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp?M=IP" & strQsSID3)
End If



'Read in the post ID
lngPostID = LngC(Request.Form("PID"))

'Read in the forum ID that the user wants to move the post to
lngMoveToForumID = LngC(Request.Form("forum"))


'Query the datbase to get the forum ID for this post
strSQL = "SELECT " & strDbTable & "Topic.Forum_ID " & _
"FROM " & strDbTable & "Topic" & strDBNoLock & ", " & strDbTable & "Thread" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Topic.Topic_ID = " & strDbTable & "Thread.Topic_ID AND " & strDbTable & "Thread.Thread_ID = " & lngPostID & ";"

'Query the database
rsCommon.Open strSQL, adoCon


'If there is a record returened read in the forum ID
If NOT rsCommon.EOF Then
	intForumID = CInt(rsCommon("Forum_ID"))
End If

'Clean up
rsCommon.Close


'Call the moderator function and see if the user is a moderator
If blnAdmin = False Then blnModerator = isModerator(intForumID, intGroupID)


'If the user is not a moderator or admin then keck em
If blnAdmin = false AND  blnModerator = false Then
	
	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If


'Read in the category name from the database
'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "" & _
"SELECT "
If strDatabaseType = "SQLServer" OR strDatabaseType = "Access" Then
	strSQL = strSQL & " TOP 400 "
End If
strSQL = strSQL & _
" " &  strDbTable & "Topic.Topic_ID, " & strDbTable & "Topic.Subject " & _
"FROM " & strDbTable & "Topic" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Topic.Forum_ID = " & lngMoveToForumID & " " & _
"ORDER BY " & strDbTable & "Topic.Last_Thread_ID DESC "
'mySQL limit operator
If strDatabaseType = "mySQL" Then
	strSQL = strSQL & " LIMIT 400"
End If
strSQL = strSQL & ";"

'Query the database
rsCommon.Open strSQL, adoCon
		

'Place the subscribed topics into an array
If NOT rsCommon.EOF Then
	
	'Read in the row from the db using getrows for better performance
	sarryForumSubjects = rsCommon.GetRows()
End If

'Clean up
rsCommon.Close
Call closeDatabase()
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title>Discussion Forum Move Post</title>

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
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" OnLoad="self.focus();">
<table width="100%"  border="0" cellpadding="3" cellspacing="0" class="tableTopRow">
 <form method="post" name="frmMovePost" action="move_post.asp<% = strQsSID1 %>">
   <tr class="tableTopRow">
    <td colspan="2"><h1><% = strTxtMovePost %></h1></td>
   </tr>
   <tr class="tableRow">
    <td colspan="2">
     <br />
     <% = strTxtSelectTopicToMovePostTo %>
     <br />
     <br />
     <% = strTxtSelectTheTopicYouWouldLikeThisPostToBeIn %>
     <br />
     <select name="topicSelect"><%

'If there are records in the array display them
If isArray(sarryForumSubjects) Then


	'Loop through all the categories in the database
	Do WHILE intCurrentRecord =< UBound(sarryForumSubjects, 2)
	
		'Read in the deatils for the category
		lngTopicID = CLng(sarryForumSubjects(0,intCurrentRecord))
		strTopicSubject = sarryForumSubjects(1,intCurrentRecord)
	
		'Display a link in the link list to the forum
		Response.Write vbCrLf & "      <option value=""" & lngTopicID & """>" & strTopicSubject & "</option>"
	
		'Move to the next record
		intCurrentRecord = intCurrentRecord + 1
	Loop
End If


%>
     </select>
     <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
     <input type="hidden" name="PID" id="PID" value="<% = lngPostID %>" />
     <input type="hidden" name="toFID" id="toFID" value="<% = lngMoveToForumID %>" />
     <br />
     <br />
     <% = strTxtOrTypeTheSubjectOfANewTopic %>
     <br />
     <input type="text" name="subject" id="subject" size="30" maxlength="41" />
     <br />
     <br />
    </td>
   </tr>
   <tr class="tableBottomRow">
    <td width="38%" valign="top"><%

'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnLCode Then
	Response.Write("<span class=""text"" style=""font-size:10px"">Forum Software by <a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank""  style=""font-size:10px"">Web Wiz Forums&reg;</a> version " & strVersion & "</span>")
End If 
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******      
      
      %></td>
    <td width="24%" align="right"><input type="submit" name="Submit" id="Submit" value="<% = strTxtMovePost %>">&nbsp;<input type="button" name="cancel" value=" <% = strTxtCancel %> " onclick="window.close()">        
     <input type="hidden" name="postBack" value="true"></td>
   </tr>
  </form>
</table>
</body>
</html>


