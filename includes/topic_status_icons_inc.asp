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




Dim strTopicIconSrc
Dim strTopicIconBgSrc
Dim strTopicIconTitle

strTopicIconSrc = "topic"
strTopicIconBgSrc = ""
strTopicIconTitle = ""




'If a hot topic
If lngNumberOfReplies >= intNumHotReplies OR lngNumberOfViews >= intNumHotViews Then 
	strTopicIconSrc = strTopicIconSrc & "_hot"
	strTopicIconTitle = strTopicIconTitle & strTxtHot & " "
End If

	
'If a locked topic
If blnTopicLocked Then 
	strTopicIconSrc = strTopicIconSrc & "_locked"
	strTopicIconTitle = strTopicIconTitle & strTxtLocked & " "
End If


'If a sticky topic
If  intPriority = 1 Then 
	strTopicIconSrc = strTopicIconSrc & "_sticky"
	strTopicIconTitle = strTopicIconTitle & strTxtSticky & " "
End If






'Hidden topics
If blnHideTopic Then 
	strTopicIconBgSrc = "topic_hidden"
	strTopicIconTitle = strTxtHiddenTopic

'Announcements	
ElseIf intPriority => 2 Then
	strTopicIconBgSrc = "announcement"
	strTopicIconTitle = strTopicIconTitle & strTxtHighPriorityPost

'Moved
ElseIf intMovedForumID = intForumID AND intMovedForumID <> 0 Then
	strTopicIconBgSrc = "moved"
	strTopicIconTitle = strTopicIconTitle & strTxtMoved & " " & strTxtTopic

'Events
ElseIf isDate(dtmEventDate) Then
	strTopicIconBgSrc = "event"
	strTopicIconTitle = strTopicIconTitle & strTxtEvent & ": " & stdDateFormat(dtmEventDate, False)
	If isDate(dtmEventDateEnd) Then strTopicIconTitle = strTopicIconTitle & " - " & stdDateFormat(dtmEventDateEnd, False)

'Polls
ElseIf lngPollID > 0 Then 
	strTopicIconBgSrc = "poll"
	strTopicIconTitle = strTopicIconTitle & strTxtPoll2

'Normal Topic
Else
	strTopicIconBgSrc = "topic"
	strTopicIconTitle = strTopicIconTitle & strTxtTopic
End If




'If unread posts
If intUnReadPostCount = 1 Then
	strTopicIconSrc = strTopicIconSrc & "_new"
	strTopicIconTitle = strTopicIconTitle & " [1 " & strTxtNewPost & "]"
ElseIf intUnReadPostCount > 1 Then
	strTopicIconSrc = strTopicIconSrc & "_new"
	strTopicIconTitle = strTopicIconTitle & " [" & intUnReadPostCount & " " & strTxtNewPosts & "]"
End If



'If there is no extra icons to display with the topic overlay it with a blank image
If strTopicIconSrc = "topic" Then strTopicIconSrc = "topic_blank"
	



'Display the topic status icon
Response.Write("<div class=""topicIcon"" style=""background-image: url('" & strImagePath & strTopicIconBgSrc & "." & strForumImageType & "');"">" & _
"<img src=""" & strImagePath & strTopicIconSrc & "." & strForumImageType & """ border=""0"" alt=""" & strTopicIconTitle & """ title=""" & strTopicIconTitle & """ />" & _
"</div>")

%>