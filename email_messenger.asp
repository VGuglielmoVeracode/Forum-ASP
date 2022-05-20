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
Dim lngToUserID		'Holds the user ID of who the email is to
Dim strToUser		'Holds the user name of the person the email is to
Dim blnShowEmail	'set to true if the user allws emailing to them
Dim strToEmail		'Holds the email address of who the email is to
Dim strFromEmail	'Holds the email address of who the email is from
Dim blnEmailSent	'Set to true if the email has been sent
Dim strEmailBody
Dim strSubject
Dim strUsername
Dim strPassword
Dim strFormID
Dim strEmailMessageBody
Dim strEmailSubject



blnSslEnabledPage = True


'Get who the email is to
lngToUserID = LngC(Request("SEID"))

'If there is no recopinet for the email then send em to homepage
If Request("SEID") = "" OR blnEmailMessenger = False Then
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


'Initlise variables
blnEmailSent = False

'Get the email address and name of the person the email is to be sent to

'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT " & strDbTable & "Author.Username, " & strDbTable & "Author.Author_email, " & strDbTable & "Author.Show_email " & _
"FROM " & strDbTable & "Author " & strDBNoLock & " " & _
"WHERE " & strDbTable & "Author.Author_ID = " & lngToUserID

'Query the database
rsCommon.Open strSQL, adoCon

'Read in the details from the user
If NOT rsCommon.EOF Then

	strToUser = rsCommon("Username")
	strToEmail = rsCommon("Author_email")
	blnShowEmail = CBool(rsCommon("Show_email"))
End If

'Clean up
rsCommon.Close


'Get the email of who the email is from
'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT " & strDbTable & "Author.Author_email " & _
"FROM " & strDbTable & "Author " & strDBNoLock & " " & _
"WHERE " & strDbTable & "Author.Author_ID = " & lngLoggedInUserID

'Query the database
rsCommon.Open strSQL, adoCon

'Read in the details from the user
If NOT rsCommon.EOF Then

	strFromEmail = rsCommon("Author_email")
End If

'Clean up
rsCommon.Close


'If this is a post back send the mail
If Request.Form("postBack") Then
	
	'Check the session ID to stop spammers using the email form
	Call checkFormID(Request.Form("formID"))
	
	'Read in form values
	strEmailMessageBody = Request.Form("message")
	strEmailSubject = Request.Form("subject")
	
	strEmailSubject = removeAllTags(strEmailSubject)
	
	

	'Initilalse the body of the email message
	strEmailBody = strTxtHi & " " & strToUser & "," & _
	vbCrLf & vbCrLf & strTxtTheFollowingEmailHasBeenSentToYouBy & " '" & strLoggedInUsername & "' " & strTxtFromYourAccountOnThe & strMainForumName & "." & _
	vbCrLf & vbCrLf & strTxtIfThisMessageIsAbusive & ": - " & _
	vbCrLf & vbCrLf & strForumEmailAddress & _
	vbCrLf & vbCrLf & strTxtToReplyPleaseEmailContact & " '" & strLoggedInUsername & "' " & strTxtThroughTheirForumProfileAtLinkBelow & ":-" & _
	vbCrLf & vbCrLf & strForumPath & "member_profile.asp?PF=" & lngLoggedInUserID & _
	vbCrLf & vbCrLf & strTxtMessageSent & ": -" & _
	vbCrLf & "---------------------------------------------------------------------------------------" & _
	vbCrLf & vbCrLf & strEmailMessageBody

	'Inititlaise the subject of the e-mail
	strSubject = strEmailSubject

	'Send the e-mail using the Send Mail function created on the send_mail_function.inc file
	blnEmailSent = SendMail(strEmailBody, strToUser, decodeString(strToEmail), strWebsiteName, decodeString(strForumEmailAddress), strSubject, strMailComponent, false)

	'If the user wants a copy of the email as well send em one
	If Request.Form("mySelf") Then
		Call SendMail(strEmailBody, strLoggedInUsername, decodeString(strFromEmail), strWebsiteName, decodeString(strForumEmailAddress), strSubject, strMailComponent, false)
	End If

End If


'If active users is enabled update the active users application array
If blnActiveUsers Then
	'Call active users function
	saryActiveUsers = activeUsers(strTxtEmailMessenger, "", "", 0)
End If




'Create a form ID
strFormID = getSessionItem("KEY")


'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtEmailMessenger

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title><% = strTxtEmailMessenger %></title>
<meta name="generator" content="Web Wiz Forums" />

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

<script language="JavaScript">

//Function to check form is filled in correctly before submitting
function CheckForm () {

	var errorMsg = "";
	var formArea = document.getElementById('frmEmailMsg');

	//Check for a subject
	if (formArea.subject.value==""){
		errorMsg += "\n<% = strTxtErrorTopicSubject %>";
	}

	//Check for message
	if (formArea.message.value==""){
		errorMsg += "\n<% = strTxtNoMessageError %>";
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

	//Disable submit button
	document.getElementById('Submit').disabled=true;
	
	return true;
}
</script>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
<!-- #include file="includes/header.asp" --> 
<!-- #include file="includes/status_bar_header_inc.asp" --> 
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% = strTxtEmailMessenger %></h1></td> 
 </tr> 
</table>
<br /><%
    
'If an error has occured display message
If blnEmailSent = False AND strEmailErrorMessage <> "" Then    	
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

'If the email is sent diaply a message
ElseIf blnEmailSent Then
	
	%>
<table cellspacing="1" cellpadding="3" class="tableBorder" align="center">
  <tr class="tableLedger">
   <td><% = strTxtEmailMessenger %></td>
  </tr>
  <tr class="tableRow"> 
   <td align="center"><br /><% = strTxtYourEmailHasBeenSentTo & " " & strToUser %><br /><br /></td>
 </tr>
</table><br /><%

'Else the email is not sent
Else

	'If the users account is suspended then let them know
	If blnActiveMember = False OR blnBanned Then
		%>
<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
  <tr>
    <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtError %></strong></td>
  </tr>
  <tr>
    <td><%
		'If mem suspended display message
		If blnBanned Then
			Response.Write(strTxtForumMemberSuspended)
		'Else account not yet active
		Else
			Response.Write("<br /><br />" & strTxtForumMembershipNotAct)
			If blnMemberApprove = False Then Response.Write("<br /><br />" & strTxtToActivateYourForumMem)
		
			'If admin activation is enabled let the user know
			If blnMemberApprove Then
				Response.Write("<br /><br />" & strTxtYouAdminNeedsToActivateYourMembership)
			'If email is on then place a re-send activation email link
			ElseIf blnEmailActivation AND blnLoggedInUserEmail Then 
				Response.Write("<br /><br /><a href=""javascript:winOpener('resend_email_activation.asp" & strQsSID1 & "','actMail',1,1,475,300)"">" & strTxtResendActivationEmail & "</a>")
			End If
		End If
		
	%></td>
  </tr>
</table><br /><%
	
	'Else if the user is not logged in so let them know to login
	ElseIf intGroupID = 2 Then
	
	%>
<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
  <tr>
    <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtError %></strong></td>
  </tr>
  <tr>
    <td><% = strTxtMustBeRegistered %></td>
  </tr>
</table>
	<!--#include file="includes/login_form_inc.asp" --><%

	'Else If the to user doesn't have an email address then the user can not send to them
	ElseIf isNull(strToEmail) Or strToEmail = "" Then
	
	%>
<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
  <tr>
    <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtError %></strong></td>
  </tr>
  <tr>
    <td><% = strTxtYouCanNotEmail & " " & strToUser & ", " & strTxtTheyDontHaveAValidEmailAddr %></td>
  </tr>
</table><br /><%

	'Else If the current user doesn't have a valid email address in their profile then they can't send an email
	ElseIf isNull(strFromEmail) OR strFromEmail = "" Then
		
	%>
<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
  <tr>
    <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtError %></strong></td>
  </tr>
  <tr>
    <td><% = strTxtYouCanNotEmail & " " & strToUser & ", " & strTxtYouDontHaveAValidEmailAddr %></td>
  </tr>
</table><br /><%

	'Else If the to user has choosen to hide their email address
	ElseIf blnShowEmail = False AND blnAdmin = False Then
		
	%>
<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
  <tr>
    <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtError %></strong></td>
  </tr>
  <tr>
    <td><% = strTxtYouCanNotEmail & " " & strToUser & ", " & strTxtTheyHaveChoosenToHideThierEmailAddr %></td>
  </tr>
</table><br /><%

	'Else show the form so the person can be emailed
	Else

%>
<form method="post" name="frmEmailMsg" id="frmEmailMsg" action="email_messenger.asp<% = strQsSID1 %>" onSubmit="return CheckForm();" onReset="return confirm('<% = strResetFormConfirm %>');"> 
 <table cellspacing="1" cellpadding="3" class="tableBorder" align="center">
    <tr class="tableLedger">
         <td colspan="2"><% = strTxtEmailMessenger %></td>
     </tr>
     <tr class="tableRow"> 
         <td colspan="2">*<% = strTxtRequiredFields %></td> 
        </tr> 
        <tr class="tableRow"> 
         <td align="right" width="15%"><% = strTxtRecipient %>:</td> 
         <td width="70%"><% = strToUser %></td> 
        </tr> 
        <tr class="tableRow"> 
         <td align="right" width="15%"><% = strTxtSubjectFolder %>*:</td> 
         <td width="70%"> <input type="text" name="subject" size="30" maxlength="41" tabindex="1"></td> 
        </tr> 
        <tr class="tableRow"> 
         <td valign="top" align="right" width="15%"><% = strTxtMessage %>*:<br /> 
          <br /> 
          <span class="smText"><% = strTxtNoHTMLorForumCodeInEmailBody %></span></td> 
         <td width="70%" valign="top"><textarea name="message" cols="57" rows="12" tabindex="2"></textarea></td> 
        <tr class="tableRow"> 
         <td align="right" width="15%">&nbsp;</td> 
         <td width="70%">&nbsp;<label><input type="checkbox" name="mySelf" value="True" tabindex="3" /><% = strTxtSendACopyOfThisEmailToMyself %></label></td> 
        </tr> 
        <tr class="tableBottomRow">
        <td colspan="2" align="center"><input name="SEID" type="hidden" id="to" value="<% = lngToUserID %>"> 
          <input name="postBack" type="hidden" id="postBack" value="true">
          <input type="hidden" name="formID" id="formID" value="" /> 
          <input type="submit" name="Submit" id="Submit" value="<% = strTxtSendEmail %>" tabindex="4" /> 
          <input type="reset" name="Reset" id="Reset" value="<% = strTxtResetForm %>" tabindex="5" /> 
         </td> 
        </tr> 
 </table> 
</form><%
	End If
End If

'Clean up
Call closeDatabase()
%>  
<div align="center"> 
<%
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

'Display the process time
If blnShowProcessTime Then Response.Write "<span class=""smText""><br /><br />" & strTxtThisPageWasGeneratedIn & " " & FormatNumber(Timer() - dblStartTime, 3) & " " & strTxtSeconds & "</span>"
%> 
</div>
<!-- #include file="includes/footer.asp" -->
