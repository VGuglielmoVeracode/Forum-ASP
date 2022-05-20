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

'If included in an SSL page then the IFrame also needs to be SSL
If BoolC(Request.QueryString("SSL")) Then blnSslEnabledPage = True



'Get the forum ID
intForumID = LngC(Request.QueryString("FID"))

'Clean up
Call closeDatabase()

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Search</title>
<meta name="generator" content="Web Wiz Forums" />
<meta name="robots" content="noindex, follow" />
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
</head>
<body class="dropDownSearch" style="border-width: 0px;visibility: visible;margin:4px;">
<form action="<% If intGroupID = 2 Then Response.Write("search_form.asp") Else Response.Write("search_process.asp") %><% = strQsSID1 %>" method="post" name="dropDownSearch" target="_parent" id="dropDownSearch">
 <div>
  <strong><% = strTxtSearchTheForum %></strong>
 </div>
 <div>
  <span style="line-height: 9px;"><br /></span>
  <input name="KW" id="KW" type="text" maxlength="35" style="width: 160px;" />
  <input type="submit" name="Submit" value="<% = strTxtGo %>" />
 </div>
 <div class="smText">
  <input name="resultType" type="radio" value="posts" checked="checked" />
  <% = strTxtShowPosts %>
  &nbsp;&nbsp;&nbsp;&nbsp;
  <input name="resultType" type="radio" value="topics" />
  <% = strTxtShowTopics %>
  <input name="AGE" type="hidden" id="AGE" value="<% = intSearchTimeDefault %>" />
  <input name="searchIn" type="hidden" id="searchIn" value="body" />
  <input name="DIR" type="hidden" id="DIR" value="newer" />
  <input name="forumID" type="hidden" id="forumID" value="<% = intForumID %>" />
  <input name="FID" type="hidden" id="FID" value="<% = intForumID %>" />
 </div>
 <div>
  <span style="line-height: 8px;"><br /></span>
  <a href="search_form.asp?FID=<% = intForumID & strQsSID2 %>" target="_parent" class="smLink"><% = strTxtAdvancedSearch %></a>
 </div>
</form>
</body>
</html>