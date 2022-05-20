<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="functions/functions_date_time_format.asp" -->
<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz Forums
'**  http://www.webwizforums.com
'**                                                              
'**  Copyright ©2001-2019 Web Wiz Ltd. All rights reserved.   
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


Response.ContentType = "text/html"


'Clean up
Call closeDatabase()



'Make sure this page is not cached
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"


Dim lngTopicID
Dim strCanonicalURL
Dim strSubject


'Read in topic share details
lngTopicID = LngC(Request.QueryString("TID"))
strCanonicalURL = Trim(Mid(Request.QueryString("URL"), 1, 150))
strSubject = Trim(Mid(Request.QueryString("Title"), 1, 80))


Response.Write(" " & strTxtShareThisPageOnTheseSites)

Response.Write("<div align=""center"">")
'Email Topic Option
If intGroupID <> 2 AND blnEmail AND blnActiveMember Then Response.Write("<a href=""javascript:winOpener('email_topic.asp?TID=" & lngTopicID & strQsSID2 & "','email_friend',0,1,440,480)""><img src=""" & strImagePath & "social_email_friend.png"" alt=""" & strTxtEmailTopic & """ title=""" & strTxtEmailTopic & """ onmouseover=""fadeImage(this)"" onmouseout=""unFadeImage(this)"" vspace=""5"" hspace=""5"" width=""32"" height=""32"" /></a>")
'Social share links
Response.Write(_
"<a href=""http://www.blinklist.com/index.php?Action=Blink/addblink.php&Url=" & Server.URLEncode(strCanonicalURL) & "&Title=" & Server.URLEncode(strSubject) & """ target=""_blank""><img src=""" & strImagePath & "social_blinklist.png"" alt=""" & strTxtPostThisTopicTo & " Blinklist""  title=""" & strTxtPostThisTopicTo & " Blinklist"" onmouseover=""fadeImage(this)"" onmouseout=""unFadeImage(this)"" vspace=""5"" hspace=""5"" width=""32"" height=""32"" /></a>" & _
"<a href=""http://del.icio.us/post?url=" & Server.URLEncode(strCanonicalURL) & "&title=" & Server.URLEncode(strSubject) & """ target=""_blank""><img src=""" & strImagePath & "social_delicious.png"" alt=""" & strTxtPostThisTopicTo & " Delicious""  title=""" & strTxtPostThisTopicTo & " Delicious"" onmouseover=""fadeImage(this)"" onmouseout=""unFadeImage(this)"" vspace=""5"" hspace=""5"" width=""32"" height=""32"" /></a>" & _
"<a href=""http://digg.com/submit?url=" & Server.URLEncode(strCanonicalURL) & "&title=" & Server.URLEncode(strSubject) & """ target=""_blank""><img src=""" & strImagePath & "social_digg.png"" alt=""" & strTxtPostThisTopicTo & " Digg"" title=""" & strTxtPostThisTopicTo & " Digg"" onmouseover=""fadeImage(this)"" onmouseout=""unFadeImage(this)"" vspace=""5"" hspace=""5"" width=""32"" height=""32"" /></a>" & _
"<a href=""http://www.facebook.com/share.php?u=" & Server.URLEncode(strCanonicalURL) & "&title=" & Server.URLEncode(strSubject) & """ target=""_blank""><img src=""" & strImagePath & "social_facebook.png"" alt=""" & strTxtPostThisTopicTo & " Facebook"" title=""" & strTxtPostThisTopicTo & " Facebook"" onmouseover=""fadeImage(this)"" onmouseout=""unFadeImage(this)"" vspace=""5"" hspace=""5"" width=""32"" height=""32"" /></a>" & _
"<a href=""http://furl.net/storeIt.jsp?u=" & Server.URLEncode(strCanonicalURL) & "&t=" & Server.URLEncode(strSubject) & """ target=""_blank""><img src=""" & strImagePath & "social_furl.png"" alt=""" & strTxtPostThisTopicTo & " Furl"" title=""" & strTxtPostThisTopicTo & " Furl"" onmouseover=""fadeImage(this)"" onmouseout=""unFadeImage(this)"" vspace=""5"" hspace=""5"" width=""32"" height=""32"" /></a>" & _
"<a href=""http://www.google.com/bookmarks/mark?op=edit&output=popup&bkmk=" & Server.URLEncode(strCanonicalURL) & "&title=" & Server.URLEncode(strSubject) & """ target=""_blank""><img src=""" & strImagePath & "social_google_bookmarks.png"" alt=""" & strTxtPostThisTopicTo & " Google Boomarks"" title=""" & strTxtPostThisTopicTo & " Google Boomarks"" onmouseover=""fadeImage(this)"" onmouseout=""unFadeImage(this)"" vspace=""5"" hspace=""5"" width=""32"" height=""32"" /></a>" & _
"<a href=""http://www.linkedin.com/shareArticle?url=" & Server.URLEncode(strCanonicalURL) & "&title=" & Server.URLEncode(strSubject) & """ target=""_blank""><img src=""" & strImagePath & "social_linkedin.png"" alt=""" & strTxtPostThisTopicTo & " LinkedIn"" title=""" & strTxtPostThisTopicTo & " LinkedIn"" onmouseover=""fadeImage(this)"" onmouseout=""unFadeImage(this)"" vspace=""5"" hspace=""5"" width=""32"" height=""32"" /></a>" & _
"<a href=""http://www.myspace.com/Modules/PostTo/Pages/?l=3&u=" & Server.URLEncode(strCanonicalURL) & "&t=" & Server.URLEncode(strSubject) & """ target=""_blank""><img src=""" & strImagePath & "social_myspace.png"" alt=""" & strTxtPostThisTopicTo & " MySpace"" title=""" & strTxtPostThisTopicTo & " MySpace"" onmouseover=""fadeImage(this)"" onmouseout=""unFadeImage(this)"" vspace=""5"" hspace=""5"" width=""32"" height=""32"" /></a>" & _
"<a href=""http://www.newsvine.com/_wine/save?u=" & Server.URLEncode(strCanonicalURL) & "&h=" & Server.URLEncode(strSubject) & """ target=""_blank""><img src=""" & strImagePath & "social_newsvine.png"" alt=""" & strTxtPostThisTopicTo & " Newsvine"" title=""" & strTxtPostThisTopicTo & " Newsvine"" onmouseover=""fadeImage(this)"" onmouseout=""unFadeImage(this)"" vspace=""5"" hspace=""5"" width=""32"" height=""32"" /></a>" & _
"<a href=""http://www.reddit.com/submit?url=" & Server.URLEncode(strCanonicalURL) & "&title=" & Server.URLEncode(strSubject) & """ target=""_blank""><img src=""" & strImagePath & "social_reddit.png"" alt=""" & strTxtPostThisTopicTo & " reddit"" title=""" & strTxtPostThisTopicTo & " reddit"" onmouseover=""fadeImage(this)"" onmouseout=""unFadeImage(this)"" vspace=""5"" hspace=""5"" width=""32"" height=""32"" /></a>" & _
"<a href=""http://www.stumbleupon.com/submit?url=" & Server.URLEncode(strCanonicalURL) & "&title=" & Server.URLEncode(strSubject) & """ target=""_blank""><img src=""" & strImagePath & "social_stumbleupon.png"" alt=""" & strTxtPostThisTopicTo & " StumbleUpon"" title=""" & strTxtPostThisTopicTo & " StumbleUpon"" onmouseover=""fadeImage(this)"" onmouseout=""unFadeImage(this)"" vspace=""5"" hspace=""5"" width=""32"" height=""32"" /></a>" & _
"<a href=""http://technorati.com/faves?add=" & Server.URLEncode(strCanonicalURL) & """ target=""_blank""><img src=""" & strImagePath & "social_technorati.png"" alt=""" & strTxtPostThisTopicTo & " Technorati"" title=""" & strTxtPostThisTopicTo & " Technorati"" onmouseover=""fadeImage(this)"" onmouseout=""unFadeImage(this)"" vspace=""5"" hspace=""5"" width=""32"" height=""32"" /></a>" & _
"<a href=""http://twitter.com/home?status=" & Server.URLEncode(strSubject) & ": " & Server.URLEncode(strCanonicalURL) & """ target=""_blank""><img src=""" & strImagePath & "social_twitter.png"" alt=""" & strTxtPostThisTopicTo & " Twitter"" title=""" & strTxtPostThisTopicTo & " Twitter"" onmouseover=""fadeImage(this)"" onmouseout=""unFadeImage(this)"" vspace=""5"" hspace=""5"" width=""32"" height=""32"" /></a>" & _
"<a href=""http://favorites.live.com/quickadd.aspx?url=" & Server.URLEncode(strCanonicalURL) & "&title=" & Server.URLEncode(strSubject) & """ target=""_blank""><img src=""" & strImagePath & "social_windows_live.png"" alt=""" & strTxtPostThisTopicTo & " Windows Live"" title=""" & strTxtPostThisTopicTo & " Windows Live"" onmouseover=""fadeImage(this)"" onmouseout=""unFadeImage(this)"" vspace=""5"" hspace=""5"" width=""32"" height=""32"" /></a>" & _
"<a href=""http://bookmarks.yahoo.com/toolbar/savebm?u=" & Server.URLEncode(strCanonicalURL) & "&t=" & Server.URLEncode(strSubject) & """ target=""_blank""><img src=""" & strImagePath & "social_yahoo.png"" alt=""" & strTxtPostThisTopicTo & " Yahoo Bookmarks"" title=""" & strTxtPostThisTopicTo & " Yahoo Bookmarks"" onmouseover=""fadeImage(this)"" onmouseout=""unFadeImage(this)"" vspace=""5"" hspace=""5"" width=""32"" height=""32"" /></a>")

Response.Write("</div>")
%>