<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="language_files/pm_language_file_inc.asp" -->
<!--#include file="functions/functions_date_time_format.asp" -->
<!--#include file="functions/functions_send_mail.asp" -->
<!--#include file="functions/functions_format_post.asp" -->
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





'Set the buffer to true
Response.Buffer = True

'Make sure this page is not cached
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"

'Declare variables
Dim lngPmMessageID		'Private message id
Dim strPmSubject 		'Holds the subject of the private message
Dim strUsername 		'Holds the Username of the thread
Dim strEmailBody		'Holds the body of the e-mail message
Dim blnEmailSent		'set to true if an e-mail is sent
Dim strPrivateMessage		'Holds the private message


'Raed in the pm mesage number to display
lngPmMessageID = LngC(Request.QueryString("ID"))

'If Priavte messages are not on then send them away
If blnPrivateMessages = False Then 
	
	'Clean up
	Call closeDatabase()
	
	Response.Redirect("default.asp" & strQsSID1)
End If



'If the user is not allowed then send them away
If intGroupID = 2 OR blnActiveMember = False OR blnBanned Then 
	
	'Clean up
	Call closeDatabase()
	
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If



	
'Initlise the sql statement
strSQL = "SELECT " & strDbTable & "PMMessage.*, " & strDbTable & "Author.Username " & _
"FROM " & strDbTable & "Author" & strDBNoLock & ", " & strDbTable & "PMMessage" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Author.Author_ID = " & strDbTable & "PMMessage.From_ID AND " & strDbTable & "PMMessage.PM_ID = " & lngPmMessageID & " "
'If this is a link from the out box then check the from author ID to check the user can view the message
If Request.QueryString("M") = "OB" Then
	strSQL = strSQL & " AND " & strDbTable & "PMMessage.From_ID = " & lngLoggedInUserID & ";"
'Else use the to author ID to check the user can view the message
Else
	strSQL = strSQL & " AND " & strDbTable & "PMMessage.Author_ID = " & lngLoggedInUserID & ";"
End If

'Query the database
rsCommon.Open strSQL, adoCon



'If a mesage is found then send a mail
If blnLoggedInUserEmail AND blnEmail AND NOT rsCommon.EOF Then 
	
	'Read in some of the details
	strPmSubject = rsCommon("PM_Tittle")
	strUsername = rsCommon("Username")
	strPrivateMessage = rsCommon("PM_Message")
	
	'Change	the path to the	emotion	symbols	to include the path to the images
	strPrivateMessage = Replace(strPrivateMessage, "src=""smileys/smiley", "src=""" & strForumPath & "smileys/smiley", 1, -1, 1)
	
	'Initailise the e-mail body variable with the body of the e-mail
	strEmailBody = strTxtHi & " " & decodeString(strLoggedInUsername) & "," & _
	"<br /><br />" & strTxtEmailBelowPrivateEmailThatYouRequested & ":-" & _
	"<br /><br /><hr />" & _
	"<br /><b>" & strTxtPrivateMessage & " :</b> " & strPmSubject & _
	"<br /><b>" & strTxtSentBy & " :</b> " & decodeString(strUsername) & _
	"<br /><b>" & strTxtSent & " :</b> " & DateFormat(rsCommon("PM_Message_date")) & " at " & TimeFormat(rsCommon("PM_Message_date")) & "<br /><br />" & _
	strPrivateMessage
		
		
	'Call the function to send the e-mail
	blnEmailSent = SendMail(strEmailBody, decodeString(strLoggedInUsername), decodeString(strLoggedInUserEmail), strWebsiteName, decodeString(strForumEmailAddress), decodeString(strPmSubject), strMailComponent, true)
End If



'Clear server objects
rsCommon.Close
Call closeDatabase()

Response.Redirect("pm_message.asp?ES=" & blnEmailSent & "&" & Request.QueryString)
%>