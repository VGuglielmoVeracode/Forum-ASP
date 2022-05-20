<% @ Language=VBScript %>
<% Option Explicit %>
<% Response.AddHeader "X-UA-Compatible","IE=EmulateIE10" %>
<!--#include file="common.asp" -->
<!--#include file="functions/functions_edit_post.asp" -->
<!--#include file="includes/emoticons_inc.asp" -->
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

'Make sure this page is not cached
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"


'Dimension variables
Dim strMode			'Holds the mode of the page
Dim lngTopicID			'Holds the Topic ID number
Dim lngMessageID		'Holds the Thread ID of the post
Dim intCatID			'Holds the cat ID
Dim strCatName			'Holds the cat name
Dim intMasterForumID		'Holds the main forum ID
Dim strMasterForumName		'Holds the main forum name
Dim strForumName		'Holds the name of the forum
Dim blnForumLocked		'Set to true if the forum is locked
Dim intTopicPriority		'Holds the priority of the topic
Dim strPostPage 		'Holds the page the form is posted to
Dim intRecordPositionPageNum	'Holds the recorset page number to show the Threads for
Dim strMessage			'Holds the post message
Dim intPollLoopCounter		'Holds the poll loop counter
Dim intIndexPosition		'Holds the index poistion in the emiticon array
Dim intNumberOfOuterLoops	'Holds the outer loop number for rows
Dim intLoop			'Holds the loop index position
Dim intInnerLoop		'Holds the inner loop number for columns
Dim lngPollEditID		'Holds the poll ID to be edited
Dim strUsername			'For login include
Dim strPassword			'For login include
Dim strUploadedFiles		'Holds the names of any files or images uploaded
Dim strTopicIcon		'Holds the topic icon
Dim intEventYear		'Holds the year of Calendar event
Dim intEventMonth		'Holds the month of Calendar event
Dim intEventDay			'Holds the day of Calendar event
Dim intEventYearEnd		'Holds the year of Calendar event
Dim intEventMonthEnd		'Holds the month of Calendar event
Dim intEventDayEnd		'Holds the day of Calendar event
Dim strFormID			'Holds the ID for the form


'If the user is user is using a banned IP redirect to an error page
If bannedIP() Then
	
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("insufficient_permission.asp?M=IP" & strQsSID3)
End If

'Intialise variables
lngTopicID = 0
lngMessageID = 0
intTopicPriority = 0
intRecordPositionPageNum = 1
strMode = "poll"


'Read in the forum and topic ID number and mode
intForumID = IntC(Request.QueryString("FID"))


'Read in the forum details inc. cat name, forum details, and permissions (also reads in the main forum name if in a sub forum, saves on db call later)

'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "" & _
"SELECT" & strDBTop1 & " " & strDbTable & "Category.Cat_ID, " & strDbTable & "Category.Cat_name, " & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Sub_ID, " & strDbTable & "Forum.Forum_name, " & strDbTable & "Forum2.Forum_name AS Main_forum, " & strDbTable & "Forum.Password, " & strDbTable & "Forum.Forum_code, " & strDbTable & "Forum.Locked, " & strDbTable & "Forum.Show_topics, " & strDbTable & "Forum.Hide, " & strDbTable & "Permissions.* " & _
"FROM " & strDbTable & "Category" & strDBNoLock & ", " & strDbTable & "Forum" & strDBNoLock & ", " & strDbTable & "Forum AS " & strDbTable & "Forum2" & strDBNoLock & ", " & strDbTable & "Permissions" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Category.Cat_ID = " & strDbTable & "Forum.Cat_ID " & _
 	"AND " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Permissions.Forum_ID " & _
 	"AND (" & strDbTable & "Forum.Sub_ID = " & strDbTable & "Forum2.Forum_ID OR (" & strDbTable & "Forum.Sub_ID = 0 AND " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Forum2.Forum_ID)) " & _
 	"AND " & strDbTable & "Forum.Forum_ID = " & intForumID & " " & _
 	"AND (" & strDbTable & "Permissions.Author_ID = " & lngLoggedInUserID & " OR " & strDbTable & "Permissions.Group_ID = " & intGroupID & ") " & _
"ORDER BY " & strDbTable & "Forum.Sub_ID, " & strDbTable & "Forum.Forum_Order, " & strDbTable & "Permissions.Author_ID DESC" & strDBLimit1 & ";"
	

'Query the database
rsCommon.Open strSQL, adoCon

'If there is a record returned by the recordset then check to see if you need a password to enter it
If NOT rsCommon.EOF Then

	'Read in forum details from the database
	intCatID = CInt(rsCommon("Cat_ID"))
	strCatName = rsCommon("Cat_name")
	strForumName = rsCommon("Forum_name")
	strMasterForumName = rsCommon("Main_forum")
	intMasterForumID = CLng(rsCommon("Sub_ID"))
	blnForumLocked = CBool(rsCommon("Locked"))
	
	'Read in the forum permissions
	blnRead = CBool(rsCommon("View_Forum"))
	blnPollCreate = CBool(rsCommon("Poll_create"))
	blnPriority = CBool(rsCommon("Priority_posts"))
	blnModerator = CBool(rsCommon("Moderate"))
	blnEvents = CBool(rsCommon("Calendar_event"))
	
	'If the user has no read writes then kick them
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
		Response.Redirect("forum_password_form.asp?FID=" & intForumID & strQsSID3)
	End If
End If

'Clean up
rsCommon.Close



'If the forum level for the user on this forum is read only set the forum to be locked
If (blnRead = False AND blnModerator = False AND blnAdmin = False) Then blnForumLocked = True
	
	
'If a login is displayed then switch to HTTPS	
If blnPollCreate = False AND intGroupID = 2 AND blnActiveMember AND blnForumLocked = false AND blnBanned = False Then blnSslEnabledPage = True 


'Use the application session to pass around what forum this user is within
If IntC(getSessionItem("FID")) <> intForumID Then Call saveSessionItem("FID", intForumID)

'get the session key
strFormID = getSessionItem("KEY")


'Set bread crumb trail
'Display the category name
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & "<a href=""default.asp?C=" & intCatID & strQsSID2 & SeoUrlTitle(strCatName, "&title=") & """>" & strCatName & "</a>" & strNavSpacer

'Display if there is a main forum to the sub forums name
If intMasterForumID <> 0 Then strBreadCrumbTrail = strBreadCrumbTrail & "<a href=""forum_topics.asp?FID=" & intMasterForumID & strQsSID2 & SeoUrlTitle(strMasterForumName, "&title=") & """>" & strMasterForumName & "</a>" & strNavSpacer

'Display forum name
If strForumName = "" Then strBreadCrumbTrail = strBreadCrumbTrail &  strTxtNoForums Else strBreadCrumbTrail = strBreadCrumbTrail & "<a href=""forum_topics.asp?FID=" & intForumID & strQsSID2 & SeoUrlTitle(strForumName, "&title=")  & """>" & strForumName & "</a>"

strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtCreateNewPoll



%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title><% = strTxtCreateNewPoll %></title>

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

<!-- Check the from is filled in correctly before submitting -->
<script  language="JavaScript">

//Function to check form is filled in correctly before submitting
function CheckForm () {
	
	var errorMsg = "";
	var formArea = document.getElementById('frmMessageForm');

<%
'If Gecko Madis API (RTE) need to strip default input from the API
If RTEenabled = "Gecko" Then Response.Write("	//For Gecko Madis API (RTE)" & vbCrLf & "	if (formArea.message.value.indexOf('<br>') > -1 && formArea.message.value.length==6) formArea.message.value = '';" & vbCrLf)


'If this is a guest posting check that they have entered their name
If blnPollCreate And lngLoggedInUserID = 2 Then
%>	
	//Check for a name
	if (formArea.Gname.value==""){
		errorMsg += "\n<% = strTxtNoNameError %>";
	}
<%
End If	

%>
	//Check for a subject
	if (formArea.subject.value==""){
		errorMsg += "\n<% = strTxtErrorTopicSubject %>";
	}
	
	//Check for poll question
	if (formArea.pollQuestion.value==""){
		errorMsg += "\n<% = strTxtErrorPollQuestion %>";
	}
	
	//Check for poll at least two poll choices
	if ((formArea.choice1.value=="") || (formArea.choice2.value=="")){
		errorMsg += "\n<% = strTxtErrorPollChoice %>";
	}
	
	//Check for message
	if (formArea.message.value==""){
		errorMsg += "\n<% = strTxtNoMessageError %>";
	}
	
	//Check session is not expired
        if (formArea.formID.value == ''){
                errorMsg += "\n<% = strTxtWarningYourSessionHasExpiredRefreshPageFormDataWillBeLost %>";
        }
	
	//If there is aproblem with the form then display an error
	if (errorMsg != ""){
		msg = "<% = strTxtErrorDisplayLine %>\n\n";
		msg += "<% = strTxtErrorDisplayLine1 %>\n";
		msg += "<% = strTxtErrorDisplayLine2 %>\n";
		msg += "<% = strTxtErrorDisplayLine %>\n\n";
		msg += "<% = strTxtErrorDisplayLine3 %>\n";
		
		errorMsg += alert(msg + errorMsg + "\n\n");
		return false;
	}
	return true;
}
</script>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% = strTxtCreateNewPoll %></h1></td>
 </tr>
</table><br /><%

 
 
'If the user has logged in and allowed to create polls then display the from to allow the user to post a new message
If blnPollCreate = True AND blnActiveMember = True AND (blnForumLocked = False OR blnAdmin = True) Then
	
	'Update active users table array
	If blnActiveUsers Then saryActiveUsers = activeUsers(strTxtCreatingNewPost, strForumName, "forum_topics.asp?FID=" & intForumID, 0)

	%>
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
 <tr class="tableLedger">
  <td align="left"><% = strTxtEditPost %></td>
 </tr>
 <tr class="tableRow">
  <td align="left"> 
   <!--#include file="includes/message_form_inc.asp" -->
  </td>
 </tr>
</table><%

'Else there is an error so show error table
Else

	'Update active users table array
	If blnActiveUsers Then saryActiveUsers = activeUsers(strTxtCreatingNewPost & " [" & strTxtAccessDenied & "]", strForumName, "forum_topics.asp?FID=" & intForumID, 0)

%>
<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
  <tr>
    <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtError %></strong></td>
  </tr>
  <tr>
    <td><%

	'If the users account is suspended then let them know
	If blnActiveMember = False OR blnBanned Then
			
		'If mem suspended display message
		If blnBanned Then
			Response.Write(strTxtForumMemberSuspended)
		'Else account not yet active
		Else
			Response.Write("<br />" & strTxtForumMembershipNotAct)
			If blnMemberApprove = False Then Response.Write("<br /><br />" & strTxtToActivateYourForumMem)
		
			'If admin activation is enabled let the user know
			If blnMemberApprove Then
				Response.Write("<br />" & strTxtYouAdminNeedsToActivateYourMembership)
			'If email is on then place a re-send activation email link
			ElseIf blnEmailActivation AND blnLoggedInUserEmail Then 
				Response.Write("<br /><br /><a href=""javascript:winOpener('resend_email_activation.asp" & strQsSID1 & "','actMail',1,1,475,300)"">" & strTxtResendActivationEmail & "</a>")
			End If
		End If
	
	'Else if the forum is locked display a message telling the user so
	ElseIf blnForumLocked = True Then
		
		Response.Write(strTxtForumLockedByAdmim)
		
		
	'Else if the user does not have permision to create polls
	ElseIf blnPollCreate = False AND intGroupID <> 2 Then
		
		Response.Write(strTxtSorryYouDoNotHavePermissionToCreatePollsForum)
	
	'Else the user is not logged in so let them know to login before they can post a message
	Else
		Response.Write(strTxtMustBeRegisteredToPost)
	End If


%></td>
  </tr>
</table><br /><%

	'If the user can needs to login display login box
	If blnPollCreate = False AND intGroupID = 2 AND blnActiveMember AND blnForumLocked = false AND blnBanned = False Then 
		%><!--#include file="includes/login_form_inc.asp" --><%
	End If

	
End If


'Clean up
Call closeDatabase()

%>
<br />
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