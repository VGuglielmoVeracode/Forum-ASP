<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
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
Dim lngTopicID			'Holds the topic ID
Dim lngPostID			'Holds the post ID
Dim intUnReadForumPostsLoop	'Loop counter
Dim strRedirectURL
Dim strSubject

lngPostID = 0


'Read in the topic ID
lngTopicID = LngC(Request.QueryString("TID"))

'Set up redirect
strRedirectURL = "forum_posts.asp?TID=" & lngTopicID
If Request.QueryString("KW") <> "" Then strRedirectURL = strRedirectURL & "&KW=" & Server.URLEncode(Request.QueryString("KW")) 
If Request.QueryString("PR") = "3" Then strRedirectURL = strRedirectURL & "&FID=" & Server.URLEncode(Request.QueryString("FID")) & "&PR=" & Server.URLEncode(Request.QueryString("PR"))


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




'Initliase the SQL query to get all the posts in this topic that are not hidden
strSQL = "SELECT" & " " & strDbTable & "Topic.Topic_ID, " & strDbTable & "Topic.Subject, " & strDbTable & "Thread.Thread_ID, " & strDbTable & "Thread.Message_date " & _
"FROM " & strDbTable & "Topic" & strDBNoLock & ", " & strDbTable & "Thread" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Topic.Topic_ID = " & strDbTable & "Thread.Topic_ID " & _
	"AND " & strDbTable & "Topic.Topic_ID = " & lngTopicID & " "
If blnAdmin = False Then
	strSQL = strSQL & _
	"AND " & strDbTable & "Topic.Hide = " & strDBFalse & " " & _
	"AND " & strDbTable & "Thread.Hide = " & strDBFalse & " "
End If
strSQL = strSQL & _
"ORDER BY " & strDbTable & "Thread.Message_date ASC;"

'Query the database
rsCommon.Open strSQL, adoCon



'Loop through the recordset to find the last unread post in this topic
Do While NOT rsCommon.EOF

	'Read in the topic details to get (read in here so we can do a redirect if a unread post is not found)
	lngPostID = CLng(rsCommon("Thread_ID"))
	strSubject = rsCommon("Subject")
	
	'Make sure we are handing an array
	If isArray(sarryUnReadPosts) AND  CDate(rsCommon("Message_date")) > dtmLastVisitDate  Then
		
		'Loop through the unread post array
		For intUnReadForumPostsLoop = 0 to UBound(sarryUnReadPosts,2)
		
			'If this post is unread then get it.
			If sarryUnReadPosts(0,intUnReadForumPostsLoop) = CLng(rsCommon("Thread_ID")) AND sarryUnReadPosts(3,intUnReadForumPostsLoop) = "1" Then	
			
				'Clean up
				rsCommon.Close
				Call closeDatabase()
				
				'Set up the URL to redirect to
				strRedirectURL = strRedirectURL & "&PID=" & lngPostID & strQsSID3 & SeoUrlTitle(strSubject, "&title=") & "#" & lngPostID 
				
				'Redirect using 301 Moved Permanently header so that search engines do not index this file
				Response.Status = "301 Moved Permanently"
				Response.AddHeader "Location", strRedirectURL
				Response.End
				
				'Exit loop
				Exit Do
			End If
		Next	
	End If
	

	'Move next record
	rsCommon.MoveNext
Loop
	

'If we didn't find an unread post, but did find the topic go to the last post in that topic
If lngPostID <> 0 Then
	
	'Clean up
	rsCommon.Close
	Call closeDatabase()
	
	
	'Set up the URL to redirect to
	strRedirectURL = strRedirectURL  & "&PID=" & lngPostID & strQsSID3 & SeoUrlTitle(strSubject, "&title=") & "#" & lngPostID 
						
	'Redirect using 301 Moved Permanently header so that search engines do not index this file
	Response.Status = "301 Moved Permanently"
	Response.AddHeader "Location", strRedirectURL
	Response.End
	
End If
	


'Clean up
rsCommon.Close
Call closeDatabase()


'If we get here there is a mistake at the users end so send 'em to the home page of the forum
Response.Status = "301 Moved Permanently"
Response.AddHeader "Location", "default.asp" & strQsSID1 
Response.End
%>