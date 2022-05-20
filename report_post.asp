<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="functions/functions_send_mail.asp" -->
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
Dim lngPostID		'Holds the post ID number
Dim lngTopicID		'Holds the topic ID number
Dim intTopicPageNum	'Holds the topic page number
Dim strPostedMessage	'Holds the posted message



Dim lngToUserID		'Holds the user ID of who the email is to
Dim strToUser		'Holds the user name of the person the email is to
Dim blnShowEmail	'set to true if the user allws emailing to them
Dim strToEmail		'Holds the email address of who the email is to
Dim strFromEmail	'Holds the email address of who the email is from
Dim blnEmailSent	'Set to true if the email has been sent
Dim strEmailBody
Dim strSubject
Dim strMessagePoster
Dim strUsername
Dim strPassword
Dim strForumName		'Holds the forum name


'Read in the details
lngPostID = LngC(Request("PID"))
lngTopicID = LngC(Request("TID"))
intTopicPageNum = IntC(Request("PN"))
intForumID = IntC(Request("FID"))


'If there is no recopinet for the email then send em to homepage
If lngPostID = 0 OR lngTopicID = 0 OR blnEmail = False OR blnActiveMember = false Then
	Call closeDatabase()

	Response.Redirect("default.asp" & strQsSID1)
End If


'If the user is user is using a banned IP redirect to an error page
If bannedIP() Then

	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp?M=IP" & strQsSID3)

End If


'If a login is displayed then switch to HTTPS
If intGroupID = 2 Then blnSslEnabledPage = True



'Initlise variables
blnEmailSent = False


'If this is a post back send the mail
If Request.Form("postBack") AND intGroupID <> 2 Then


	'Get the post to send with the email
	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "SELECT " & strDbTable & "Author.Username, " & strDbTable & "Thread.Message " & _
	"FROM " & strDbTable & "Author" & strDBNoLock & ", " & strDbTable & "Thread" & strDBNoLock & "  " & _
	"WHERE " & strDbTable & "Author.Author_ID = " & strDbTable & "Thread.Author_ID AND " & strDbTable & "Thread.Thread_ID=" & lngPostID & ";"

	'Query the database
	rsCommon.Open strSQL, adoCon

	'Read in the message to be sent with the report
	If NOT rsCommon.EOF Then
		strPostedMessage = rsCommon("Message")
		strMessagePoster = rsCommon("Username")

		'Change	the path to the	emotion	symbols	to include the path to the images
		strPostedMessage = Replace(strPostedMessage, "src=""smileys/smiley", "src=""" & strForumPath & "smileys/smiley", 1, -1, 1)
	End If

	'Clean up
	rsCommon.Close



	'Inititlaise the subject of the e-mail
	strSubject = strTxtIssueWithPostOn & " " & strMainForumName

	'Initilalse the body of the email message
	strEmailBody = strTxtHi & "," & _
	"<br /><br />" & strTxtTheFollowingReportSubmittedBy & " " & decodeString(strLoggedInUsername) & ", " & strTxtOn & " " & strMainForumName & " " & strTxtWhoHasTheFollowingIssue & " : -" & _
	"<br /><br /><hr />" & _
	"<br />" & Replace(Request.Form("report"), vbCrLf, "<br />",	1, -1, 1) & "<br /><br />" & _
	"<hr />" & _
	"<br />" & strTxtToViewThePostClickTheLink & " : -" & _
	"<br /><a href=""" & strForumPath & "forum_posts.asp?TID=" & lngTopicID & "&PN=" & intTopicPageNum & """>" & strForumPath & "forum_posts.asp?TID=" & lngTopicID & "&PN=" & intTopicPageNum & "</a>" & _
	"<br /><br /><hr /><br /><b>" & strTxtPostedBy & ":</b> " & strMessagePoster & "<br /><br />" & _
	strPostedMessage



	'Get the email address of the boards admins to send the email to
	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "SELECT " & strDbTable & "Author.Username, " & strDbTable & "Author.Author_email " & _
	"FROM " & strDbTable & "Author" & strDBNoLock & "  " & _
	"WHERE " & strDbTable & "Author.Group_ID=1 AND " & strDbTable & "Author.Author_email <> '';"


	'Query the database
	rsCommon.Open strSQL, adoCon
	
	'Send an email to the forum email address if there are no email addresses of admins in the database
	If rsCommon.EOF Then blnEmailSent = SendMail(strEmailBody, strTxtForumAdministrator, strForumEmailAddress, strLoggedInUsername, strForumEmailAddress, strSubject, strMailComponent, true)
	
	'If there are email addresses returned send email to the forum admins
	Do while not rsCommon.EOF

		'Send the e-mail using the Send Mail function created on the send_mail_function.inc file
		blnEmailSent = SendMail(strEmailBody, rsCommon("Username"), decodeString(rsCommon("Author_email")), strLoggedInUsername, strForumEmailAddress, strSubject, strMailComponent, true)

		'Move to next record
		rsCommon.MoveNext
	Loop

	'Clean up
	rsCommon.Close




	'Get the email address of the moderators to send the email to
	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "SELECT " & strDbTable & "Author.Username, " & strDbTable & "Author.Author_email " & _
	"FROM " & strDbTable & "Permissions" & strDBNoLock & ", " & strDbTable & "Author" & strDBNoLock & "  " & _
	"WHERE " & strDbTable & "Permissions.Group_ID = " & strDbTable & "Author.Group_ID AND (" & strDbTable & "Permissions.Forum_ID=" & intForumID & " AND " & strDbTable & "Permissions.Moderate=" & strDBTrue & ") AND " & strDbTable & "Author.Author_email <> '';"
	

	'Query the database
	rsCommon.Open strSQL, adoCon

	'Send an email to the moderators
	Do while not rsCommon.EOF

		'Send the e-mail using the Send Mail function created on the send_mail_function.inc file
		blnEmailSent = SendMail(strEmailBody, rsCommon("Username"), decodeString(rsCommon("Author_email")), strLoggedInUsername, strForumEmailAddress, strSubject, strMailComponent, true)

		'Move to next record
		rsCommon.MoveNext
	Loop

	'Clean up
	rsCommon.Close
End If


'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtReportPost

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title><% = strTxtReportPost %></title>
<meta name="generator" content="Web Wiz Forums" />
<meta name="robots" content="noindex, follow" />

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
  <td align="left"><h1><% = strTxtReportPost %></h1></td>
 </tr>
</table>
<br /><%


'If the users account is suspended then let them know
If blnBanned Then
	
	%>
<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
  <tr>
    <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtError %></strong></td>
  </tr>
  <tr>
    <td><% = strTxtForumMemberSuspended %></td>
  </tr>
</table><%

'Else if the user is not logged in so let them know to login
ElseIf intGroupID = 2 Then

	%>
<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
  <tr>
    <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtError %></strong></td>
  </tr>
  <tr>
    <td><% = strTxtMustBeRegistered %><br /><br /><a href="javascript:history.back(1)"><% = strTxtReturnForumTopic %></a></td>
  </tr>
</table>
<br />
<!--#include file="includes/login_form_inc.asp" -->
<br /><%
    
'If an error has occured display message
ElseIf blnEmailSent = False AND strEmailErrorMessage <> "" Then    	
%>
<table cellspacing="1" cellpadding="3" class="errorTable" align="center">
  <tr>
    <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtError %></strong></td>
  </tr>
  <tr>
   <td>
    <% = strTxtTheEmailFailedToSendPleaseContactAdmin %>
    <br /><br /><strong>Server Error Message:-</strong>
    <br /><% = strEmailErrorMessage %>
    <br />
  </td>
 </tr>
</table><br /><%

'Else show the form so the person can be emailed
Else

%>
<form method="post" name="frmReport" id="frmReport" action="report_post.asp<% = strQsSID1 %>" onReset="return confirm('<% = strResetFormConfirm %>');">
<table cellspacing="1" cellpadding="3" class="tableBorder" align="center">
 <tr class="tableLedger">
  <td colspan="2"><% = strTxtReportPost %></td>
 </tr>
 <tr class="tableRow"><%
 
	'If the email has been sent display the appropriate message
	If blnEmailSent Then

	Response.Write(vbCrLf & "   <td><br />" & strTxtYourReportEmailHasBeenSent & "<br /><br /><a href=""default.asp"">" & strTxtReturnToDiscussionForum & "</a><br /></td>")
	Else
%>
   <td colspan="2" class="text"><% = strTxtPleaseStateProblemWithPost %><br /><br /></td>
  </tr>
  <tr  class="tableRow">
   <td valign="top" align="right" width="30%"><% = strTxtProblemWithPost %>*:</td>
   <td width="70%" valign="top"><textarea name="report" cols="70" rows="12"></textarea></td>
  </tr>
  <tr class="tableBottomRow">
   <td colspan="2" align="center">
    <input name="PID" type="hidden" id="PID" value="<% = lngPostID %>" />
    <input name="FID" type="hidden" id="FID" value="<% = intForumID %>" />
    <input name="TID" type="hidden" id="TID" value="<% = lngTopicID %>">
    <input name="PN" type="hidden" id="PN" value="<% = intTopicPageNum %>" />
    <input name="postBack" id="postBack" type="hidden" value="true" />
    <input type="submit" name="Submit" id="Submit" value="<% = strTxtSendReport %>" />
    <input type="reset" name="Reset" id="Reset" value="<% = strTxtResetForm %>" />
   </td><%

	End If

%>
  </tr>
</table>
</form><%

End If

'Clean up
Call closeDatabase()
%>
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