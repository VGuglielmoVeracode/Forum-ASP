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






Response.Buffer = True 

'Dimension variables
Dim strErrorCode		'Holds the error code of the page

'Read in the error code
strErrorCode = Request.QueryString("mode")

'Reset Server Objects
Call closeDatabase()


'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtMessageNotPosted

%>  
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title><% = strTxtMessageNotPosted %></title>

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
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% = strTxtMessageNotPosted %></h1></td>
 </tr>
</table>
<br />
<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
  <tr>
    <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtError %>: <% = strTxtMessageNotPosted %></strong></td>
  </tr>
  <tr>
    <td><%

 
'Write the error message
If strErrorCode = "posted" Then
	Response.Write(strTxtDoublePostingIsNotPermitted) 
ElseIf strErrorCode = "spam" Then
	Response.Write(strTxtYourMessageWasRejectedByTheSpamFilters)
ElseIf strErrorCode = "noSubject" Then
	Response.Write(strTxtYourMessageNoValidSubjectHeading) 
ElseIf strErrorCode = "maxS" OR strErrorCode = "maxM" Then    
	Response.Write(strTxtSpammingIsNotPermitted & "<br />" & strTxtYouHaveExceededNumOfPostAllowed) 
ElseIf strErrorCode = "noPoll" Then
	Response.Write(strTxtYourNoValidPoll) 
ElseIf strErrorCode = "FLocked" Then
	Response.Write(strTxtThisForumIsLocked) 
ElseIf strErrorCode = "TClosed" Then
	Response.Write(strTxtThisTopicIsLocked) 
ElseIf strErrorCode = "CAPTCHA" Then
	Response.Write(strTxtSecurityCodeDidNotMatch2) 
End If        
        %><br />
     <br /><%

If Request.QueryString("TID") <> "" Then
	Response.Write("<a href=""forum_posts.asp?TID=" & LngC(Request.QueryString("TID")) & strQsSID2 & """>" & strTxtReturnToDiscussionForum & "</a>")
Else
     	Response.Write("<a href=""javascript:history.back(1)"">" & strTxtReturnToDiscussionForum & "</a>")
End If

%>
   </td>
  </tr>
</table>
<br />
<div align="center">
<% 
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