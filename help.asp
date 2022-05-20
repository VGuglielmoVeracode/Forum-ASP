<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="language_files/help_language_file_inc.asp" -->
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



'If active users is enabled update the active users application array
If blnActiveUsers Then
	'Call active users function
	saryActiveUsers = activeUsers("", strTxtForumHelp, "help.asp", 0)
End If


'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtForumHelp

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title><% = strTxtForumHelp %></title>
<meta name="generator" content="Web Wiz Forums" />
<meta name="description" content="<% = strBoardMetaDescription %>" />
<meta name="keywords" content="<% = strBoardMetaKeywords %>" />

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
<a name="top"></a>
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% = strTxtForumHelp %></h1></td>
 </tr>
</table>
<br />
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
  <tr class="tableLedger">
   <td><% = strTxtChooseAHelpTopic %></td>
  </tr>
  <tr class="tableSubLedger">
   <td><% = strTxtLoginAndRegistration %></td>
  </tr>
  <tr> 
   <td class="tableRow">
    <a href="#FAQ1"><% = strTxtWhyCantILogin %></a>
    <br />
    <a href="#FAQ2"><% = strTxtDoINeedToRegister %></a>
    <br />
    <a href="#FAQ3"><% = strTxtLostPasswords %></a>
    <br />
    <a href="#FAQ4"><% = strTxtIRegisteredInThePastButCantLogin %></a>
    <br /><br />
   </td>
  </tr>
  <tr class="tableSubLedger"> 
   <td><% = strTxtUserPreferencesAndForumSettings %></td>
  </tr>
  <tr> 
   <td class="tableRow">
    <a href="#FAQ5"><% = strTxtHowDoIChangeMyForumSettings %></a>
    <br />
    <a href="#FAQ6"><% = strTxtForumTimesAndDates %></a>
    <br />
    <a href="#FAQ7"><% = strTxtWhatDoesMyRankIndicate %></a>
    <br />
    <a href="#FAQ8"><% = strTxtCanIChangeMyRank %></a>
    <br />
    <a href="#FAQ34"><% = strTxtWhatWebBrowserCanIUseForThisForum %></a>
     <br /><br />
   </td>
  </tr>
    
  <tr class="tableSubLedger"> 
   <td><% = strTxtPostingIssues %></td>
  </tr>
  <tr> 
   <td class="tableRow">
    <a href="#FAQ9"><% = strTxtHowPostMessageInTheForum %></a>
    <br />
    <a href="#FAQ10"><% = strTxtHowDeletePosts %></a>
    <br />
    <a href="#FAQ11"><% = strTxtHowEditPosts %></a>
    <br />
    <a href="#FAQ12"><% = strTxtHowSignaturToMyPost %></a>
    <br />
    <a href="#FAQ13"><% = strTxtHowCreatePoll %></a>
    <br />
    <a href="#FAQ14"><% = strTxtWhyNotViewForum %></a>
    <br />
    <a href="#FAQ28"><% = strTxtMyPostIsHiddenOrPendingApproval %></a>
    <br />
    <a href="#FAQ15"><% = strTxtInternetExplorerWYSIWYGPosting %></a>
    <br /><br />
   </td>
  </tr>
  <tr class="tableSubLedger"> 
   <td><% = strTxtMessageFormatting %></td>
  </tr>
   <td class="tableRow">
    <a href="#FAQ16"><% = strTxtWhatForumCodes %></a>
    <br />
    <a href="#FAQ17"><% = strTxtCanIUseHTML %></a>
    <br />
    <a href="#FAQ18"><% = strTxtWhatEmoticons %></a>
    <br />
    <a href="#FAQ19"><% = strTxtCanPostImages %></a>
    <br />
    <a href="#FAQ20"><% = strTxtWhatClosedTopics %></a>
    <br /><br />
   </td>
  </tr>
  <tr class="tableSubLedger"> 
   <td><% = strTxtUsergroups %></td>
  </tr>
  <tr> 
   <td class="tableRow">
    <a href="#FAQ21"><% = strTxtWhatForumAdministrators %></a>
    <br />
    <a href="#FAQ22"><% = strTxtWhatForumModerators %></a>
    <br />
    <a href="#FAQ23"><% = strTxtWhatUsergroups %></a>
    <br /><br />
   </td>
  </tr>
  <tr class="tableSubLedger"> 
   <td><% = strTxtPrivateMessaging %></td>
  </tr>
  <tr> 
   <td class="tableRow">
    <a href="#FAQ27"><% = strTxtWhatIsPrivateMessaging %></a>
    <br />
    <a href="#FAQ24"><% = strTxtIPrivateMessages %></a>
    <br />
    <a href="#FAQ25"><% = strTxtIPrivateMessagesToSomeUsers %></a>
    <br />
    <a href="#FAQ26"><% = strTxtHowCanPreventSendingPrivateMessages %></a>
    <br /><br />
   </td>
  </tr>
  <tr class="tableSubLedger"> 
   <td><% = strTxtRSSFeeds %></td>
  </tr>
  <tr> 
   <td class="tableRow">
    <a href="#FAQ29"><% = strTxtWhatIsAnRSSFeed %></a>
    <br />
    <a href="#FAQ30"><% = strTxtHowDoISubscribeToRSSFeeds %></a>
    <br /><br />
   </td>
  </tr>
  <tr class="tableSubLedger"> 
   <td><% = strTxtCalendarSystem %></td>
  </tr>
  <tr> 
   <td class="tableRow">
    <a href="#FAQ31"><% = strTxtWhatIsTheCalendarSystem %></a>
    <br />
    <a href="#FAQ32"><% = strTxtHowDoICreateCalendarEvent %></a>
    <br /><br />
   </td>
  </tr>
  
  <tr class="tableSubLedger"> 
   <td><% = strTxtCookiePrivacyNotice %></td>
  </tr>
  <tr> 
   <td class="tableRow">
    <a href="#FAQ36"><% = strTxtWhatAreCookies %></a>
    <br />
    <a href="#FAQ35"><% = strTxtInformationOnCookiesSetByThisForum %></a>
    <br /><br />
   </td>
  </tr><%

If blnLCode Then
	
%>
  <tr class="tableSubLedger"> 
   <td><% = strTxtAbout %></td>
  </tr>
  <tr> 
   <td class="tableRow">
    <a href="#FAQ33"><% = strTxtWhatSoftwareIsUsedForThisForum %></a>
   </td>
  </tr><%

End If

%>
</table>
<br>
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
  <tr class="tableLedger">
   <td><% = strTxtLoginAndRegistration %></td>
  </tr>
  <tr class="tableSubLedger">
   <td><a name="FAQ1" id="FAQ1"></a><% = strTxtWhyCantILogin %></td>
  </tr>
  <tr class="tableRow">
   <td align="justify"><% = strTxtFAQ1 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td>
  </tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ2"></a><% = strTxtDoINeedToRegister %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ2 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td>
  </tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ3"></a><% = strTxtLostPasswords %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ3 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td>
  </tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ4"></a><% = strTxtIRegisteredInThePastButCantLogin %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ4 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a></td>
  </tr>
</table>

<br />

<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
  <tr class="tableLedger">
    <td><% = strTxtUserPreferencesAndForumSettings %></td>
  </tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ5"></a><% = strTxtHowDoIChangeMyForumSettings %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ5 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td></tr>
  <tr class="tableSubLedger">
    <td align="justify"><a name="FAQ6"></a><% = strTxtForumTimesAndDates %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ6 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td></tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ7"></a><% = strTxtWhatDoesMyRankIndicate %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ7 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td></tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ8"></a><% = strTxtCanIChangeMyRank %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ8 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td></tr>
    <tr class="tableSubLedger">
    <td><a name="FAQ34"></a><% = strTxtWhatWebBrowserCanIUseForThisForum %></td>
  </tr>
  <tr class="tableRow">
   <td align="justify"><% = strTxtFAQ34 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a>
  </td>
 </tr>
</table>

<br />

<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
  <tr class="tableLedger">
    <td><% = strTxtPostingIssues %></td>
  </tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ9"></a><% = strTxtHowPostMessageInTheForum %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ9 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td></tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ10"></a><% = strTxtHowDeletePosts %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ10 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td></tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ11"></a><% = strTxtHowEditPosts %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ11 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td></tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ12"></a><% = strTxtHowSignaturToMyPost %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ12 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td></tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ13"></a><% = strTxtHowCreatePoll %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ13 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td></tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ14"></a><% = strTxtWhyNotViewForum %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ14 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td></tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ28"></a><% = strTxtMyPostIsHiddenOrPendingApproval %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ28 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td></tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ15"></a><% = strTxtInternetExplorerWYSIWYGPosting %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ15 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a></td></tr>
</table>

<br />

<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
  <tr class="tableLedger">
    <td><% = strTxtMessageFormatting %></td>
  </tr>
  <tr class="tableSubLedger">
    <td align="justify"><a name="FAQ16"></a><% = strTxtWhatForumCodes %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ16 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td></tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ17" id="FAQ17"></a><% = strTxtCanIUseHTML %></td>
  </tr>	
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ17 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td></tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ18"></a><% = strTxtWhatEmoticons %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ18 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td></tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ19"></a><% = strTxtCanPostImages %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ19 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td></tr>
   <tr class="tableSubLedger">
    <td><a name="FAQ20"></a><% = strTxtWhatClosedTopics %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ20 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a></td></tr>
</table>

<br /> 

<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
  <tr class="tableLedger">
    <td><% = strTxtUsergroups %></td>
  </tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ21"></a><% = strTxtWhatForumAdministrators %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ21 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td>
  </tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ22"></a><% = strTxtWhatForumModerators %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ22 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td>
  </tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ23"></a><% = strTxtWhatUsergroups %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ23 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a></td>
  </tr>
</table>

<br />

<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
  <tr class="tableLedger">
    <td><% = strTxtPrivateMessaging %></td>
  </tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ27"></a><% = strTxtWhatIsPrivateMessaging %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ27 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td>
  </tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ24"></a><% = strTxtIPrivateMessages %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ24 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td>
  </tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ25"></a><% = strTxtIPrivateMessagesToSomeUsers %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ25 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td>
  </tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ26"></a><% = strTxtHowCanPreventSendingPrivateMessages %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ26 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a></td>
  </tr>
</table>

<br />

<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
  <tr class="tableLedger">
    <td><% = strTxtRSSFeeds %></td>
  </tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ29"></a><% = strTxtWhatIsAnRSSFeed %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ29 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td>
  </tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ30"></a><% = strTxtHowDoISubscribeToRSSFeeds %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ30 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a></td>
  </tr>
</table>

<br />

<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
  <tr class="tableLedger">
    <td><% = strTxtCalendarSystem %></td>
  </tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ31"></a><% = strTxtWhatIsTheCalendarSystem %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ31 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td>
  </tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ32"></a><% = strTxtHowDoICreateCalendarEvent %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ32 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a></td>
  </tr>
</table>

<br />

<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
  <tr class="tableLedger">
    <td><% = strTxtCookiePrivacyNotice %></td>
  </tr>
 <tr class="tableSubLedger">
    <td><a name="FAQ36"></a><% = strTxtWhatAreCookies %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ36 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a><br /><br /></td></tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ35"></a><% = strTxtInformationOnCookiesSetByThisForum %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ35 %>
    	<blockquote>
    	 <% = strCookiePrefix & strTxtFAQ35sID %>
    	 <br />
    	 <% = strCookiePrefix & strTxtFAQ35sLID %>
    	 <br />
    	 <% = strCookiePrefix & strTxtFAQ35lVisit %>
    	 <br />
    	 <% = strCookiePrefix & strTxtFAQ35fID %>
    	 <br />
    	 <% = strCookiePrefix & strTxtFAQ35pID %>
    	 <br />
    	 <% = strCookiePrefix & strTxtFAQ35MobileView %>
    	</blockquote>
    <a href="#top" class="smLink"><% = strTxtBackToTop %></a></td></tr>
</table>

<br /><%

If blnLCode Then
	
%>
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
  <tr class="tableLedger">
    <td><% = strTxtAbout %></td>
  </tr>
  <tr class="tableSubLedger">
    <td><a name="FAQ33"></a><% = strTxtWhatSoftwareIsUsedForThisForum %></td>
  </tr>
  <tr class="tableRow">
    <td align="justify"><% = strTxtFAQ33 %><br /><a href="#top" class="smLink"><% = strTxtBackToTop %></a></td>
  </tr>
</table>
<br /><%

End If

%>
<table class="basicTable" cellspacing="0" cellpadding="10" align="center">
  <tr>
   <td>
    <!-- #include file="includes/forum_jump_inc.asp" -->
   </td>
 </tr>
</table>
<div align="center"><br /><%

'Clean up
Call closeDatabase()

	
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