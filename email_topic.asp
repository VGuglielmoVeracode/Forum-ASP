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




Response.Buffer = True


'Dimension variables
Dim strAuthorEmail	'Holds the users e-mail address
Dim strFormMessage	'Holds the message in the form
Dim strEmailBody	'Holds the body of the e-mail
Dim blnSentEmail	'Set to true when the e-mail is sent
Dim strSubject		'Holds the subject of the e-mail
Dim strRealName		'Holds the authors real name
Dim strFormID
Dim strRecipientName	'Holds the name the email is to
Dim strRecipientEmail	'Holds the email address this email is being sent to
Dim strSenderName	'Real name of sender
Dim strEmailMessageBody	'Holds the message of the email

'Initialise variables
blnSentEmail = False


'If the user is user is using a banned IP redirect to an error page
If bannedIP() Then
	
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("insufficient_permission.asp?M=IP" & strQsSID3)

End If


'If the user has not logged in then  or the page has not been passed with a topic id the redirect to the forum start page
If Request.QueryString("TID") = "" OR blnEmail = False OR intGroupID = 2 Then 
	
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("default.asp" & strQsSID1)
End If


'Initilise the message in the form
strFormMessage = strTxtEmailFriendMessage & " " & strMainForumName & " " & strTxtAt & ": -" & _
vbCrLf & vbCrLf & strForumPath & "forum_posts.asp?TID=" & LngC(Request.QueryString("TID")) & _
vbCrLf & vbCrLf & strTxtRegards & "," & vbCrLf & strLoggedInUsername & vbCrLf


'Read in the users email address

'Initalise the strSQL variable with an SQL statement to query the database get the thread details
strSQL = "SELECT " & strDbTable & "Author.Real_name, " & strDbTable & "Author.Author_email " & _
"FROM " & strDbTable & "Author" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Author.Author_ID=" & lngLoggedInUserID & ";"

'Query the database
rsCommon.Open strSQL, adoCon

'If there is an e-mail address for the user then read it in
If NOT rsCommon.EOF Then
	'Read in authors detals from the database
	strAuthorEmail = rsCommon("Author_email")
	strRealName = rsCommon("Real_name")
End If


'If the form has been filled in then send the form
If NOT Request.Form("ToName") = "" AND NOT Request.Form("ToEmail") = "" AND NOT Request.Form("FromName") = "" AND NOT Request.Form("message") = "" Then

	'Check the session ID to stop spammers using the email form
	Call checkFormID(Request.Form("formID"))
	
	
	'Read in form values
	strRecipientName = Request.Form("ToName")
	strRecipientEmail = Request.Form("ToEmail")
	strSenderName = Request.Form("FromName")
	strEmailMessageBody = Request.Form("message")
	
	
	strRecipientName = removeAllTags(strRecipientName)
	strSenderName = removeAllTags(strSenderName)
	
	

	'Initilalse the body of the email message
	strEmailBody = strTxtHi & " " & strRecipientName & "," & _
	vbCrLf & vbCrLf & strTxtTheFollowingEmailHasBeenSentToYouBy & " '" & strSenderName & "' " & strTxtFrom & " " & strMainForumName & "." & _
	vbCrLf & vbCrLf & strTxtIfThisMessageIsAbusive & ": - " & _
	vbCrLf & vbCrLf & strForumEmailAddress & _
	vbCrLf & vbCrLf & strTxtToReplyPleaseEmailContact & " '" & strLoggedInUsername & "' " & strTxtThroughTheirForumProfileAtLinkBelow & ":-" & _
	vbCrLf & vbCrLf & strForumPath & "member_profile.asp?PF=" & lngLoggedInUserID & _
	vbCrLf & vbCrLf & strTxtMessageSent & ": -" & _
	vbCrLf & "---------------------------------------------------------------------------------------" & _
	vbCrLf & vbCrLf & strTxtHi & " " & strRecipientName & _
	vbCrLf & vbCrLf & strEmailMessageBody

	'Inititlaise the subject of the e-mail
	strSubject = strTxtInterestingForumPostOn & " " & strWebsiteName

	'Send the e-mail using the Send Mail function created on the send_mail_function.inc file
	blnSentEmail = SendMail(strEmailBody, strRecipientName, decodeString(strRecipientEmail), strWebsiteName, decodeString(strForumEmailAddress), strSubject, strMailComponent, false)
End If

'Reset server objects
rsCommon.Close
Call closeDatabase()



'Create a form ID
strFormID = getSessionItem("KEY")






%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title><% = strTxtEmailTopicToFriend %></title>

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

<script  language="JavaScript">

//Function to check form is filled in correctly before submitting
function CheckForm() {

	var errorMsg = "";
	var formArea = document.getElementById('frmEmailTopic');

	//Check for a Friends Name
	if (formArea.ToName.value == ""){
		errorMsg += "\n<% = strTxtErrorFrinedsName %>";
	}

	//Check that the friends e-mail and it is valid address is valid
	if ((formArea.ToEmail.value=="") || (formArea.ToEmail.value.length > 0 && (formArea.ToEmail.value.indexOf("@",0) == - 1 || formArea.ToEmail.value.indexOf(".",0) == - 1))) {
		errorMsg += "\n<% = strTxtErrorFriendsEmail %>";
	}

	//Check for a Users Name
	if (formArea.FromName.value==""){
		errorMsg += "\n<% = strTxtErrorYourName %>";
	}

	//Check for a Message
	if (formArea.message.value==""){
		errorMsg += "\n<% = strTxtErrorEmailMessage %>";
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

	document.getElementById('formID').value='<% = strFormID %>';
	
	return true;
}
</script>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
</head>
<body OnLoad="self.focus();">
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="center"><h1><% = strTxtEmailTopicToFriend %></h1></td>
 </tr>
</table>
<br />
<form name="frmEmailTopic" id="frmEmailTopic" method="post" action="email_topic.asp?FID=<% = IntC(Request.QueryString("FID")) %>&TID=<% = LngC(Request.QueryString("TID")) %><% = strQsSID2 %>" onSubmit="return CheckForm();" onReset="return confirm('<% = strResetFormConfirm %>');">
 <table cellspacing="1" cellpadding="3" class="tableBorder" align="center" width="350">
    <tr class="tableLedger">
     <td colspan="2"><% = strTxtEmailTopicToFriend %></td>
    <tr><%
    
'If an error has occured display message
If blnSentEmail = False AND strEmailErrorMessage <> "" Then    	
%>
     <tr class="tableRow">
      <td colspan="2" align="center">
      	<br /><% = strTxtTheEmailFailedToSendPleaseContactAdmin %>
      	<br /><br /><strong>Server Error Message:-</strong>
        <br /><% = strEmailErrorMessage %>
        <br /><br />
      </td>
     </tr><%    	
    	

'If the email has been sent then display a message saying
ElseIf blnSentEmail Then
%>
     <tr class="tableRow">
      <td colspan="2" align="center"><br /><% = strTxtFriendSentEmail %><br /><br /></td>
     </tr><%

'If the user doesn't have a valid email address then they can not send email
ElseIf isNull(strAuthorEmail) OR strAuthorEmail = "" Then
%>
     <tr class="tableRow">
      <td colspan="2" align="center"><br /><% = strTxtYouCanNotEmailTisTopicToAFriend & " " & strTxtYouDontHaveAValidEmailAddr %><br /><br /></td>
     </tr>
     <%
'Else the e-mail has not been sent so display the form
Else
%>
     <tr class="tableRow">
      <td width="115"><% = strTxtUsername %></td>
      <td width="234"><% = strLoggedInUsername %></td>
     </tr>
     <tr class="tableRow">
      <td><% = strTxtYourEmail %></td>
      <td><% = strAuthorEmail %></td>
     </tr>
     <tr class="tableRow">
      <td><% = strTxtYourName %></td>
      <td>
       <input type="text" name="FromName" size="20" maxlength="20" value="<% = strRealName %>" onFocus="FromName.value = ''" />
      </td>
     </tr>
     <tr class="tableRow">
      <td><% = strTxtFriendsName %></td>
       <td><input type="text" name="ToName" size="20" maxlength="20" /></td>
     </tr>
     <tr class="tableRow">
      <td><% = strTxtFriendsEmail %></td>
      <td><input type="text" name="ToEmail" size="20" maxlength="50" /></td>
     </tr>
     <tr class="tableRow">
      <td colspan="2"><% = strTxtMessage %>:<br />
       <textarea name="message" cols="47" rows="7" wrap="OFF"><% = strFormMessage %></textarea>
      </td>
     </tr>
     <tr class="tableBottomRow" align="center">
      <td colspan="2">
        <input type="hidden" name="formID" id="formID" value="" />
        <input type="submit" name="Submit" id="Submit" value="<% = strTxtSendEmail %>" />
        <input type="reset" name="Reset" id="Reset" value="<% = strTxtResetForm %>" /><%

End If
%>     </td>
     </tr>
    </table>
</form>
<br />
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
  <tr>
    <td align="center"><input type="button" name="ok" onclick="javascript:window.close();" value="<% = strTxtCloseWindow %>"><br />
    <br /><%
    
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnLCode = True Then
	If blnTextLinks = True Then
		Response.Write("<span class=""text"" style=""font-size:10px"">Forum Software by <a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank""  style=""font-size:10px"">Web Wiz Forums&reg;</a> version " & strVersion & "</span>")
		
	Else
  		Response.Write("<a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank"" ><img src=""webwizforums_image.asp"" border=""0"" title=""Forum Software by Web Wiz Forums&reg; version " & strVersion & """ alt=""Forum Software by Web Wiz Forums&reg; version " & strVersion& """ /></a>")
  		
	End If
	
	Response.Write("<br /><span class=""text"" style=""font-size:10px"">Copyright &copy;2001-2019 Web Wiz Ltd.</span>")
End If
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>
  </td>
 </tr>
</table>
</body>
</html>