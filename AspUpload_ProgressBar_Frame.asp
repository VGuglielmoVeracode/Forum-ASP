<% @ EnableSessionState = False
Language=VBScript
%>
<% Option Explicit %>
<!-- #include file="language_files/RTE_language_file_inc.asp" -->
<!-- #include file="functions/functions_filters.asp" -->
<%


'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz Rich Text Editor(TM)
'**  http://www.richtexteditor.org
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




Response.AddHeader "pragma","cache"
Response.AddHeader "cache-control","public"
Response.CacheControl =	"Public"
Response.Expires = -1 

Dim strPID

'Read in the PID
strPID = removeAllTags(Request.QueryString("PID"))

%>
<html>
<head>
<title><% = strTxtUploadingFiles %></title>
<style type="text/css">
<!--
html, body {
  background: ButtonFace;
  color: ButtonText;
  font: font-family: Verdana, Arial, Helvetica, sans-serif;
  font-size: 12px;
  margin: 0px;
  padding: 4px;
}
.text {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 12px;
	color: #000000;
}
-->
</style>
</head>
<body onload="self.focus();">
<table broder="0" width="100%" cellpadding="0" cellspacing="0">
  <tr>
   <td align="center" class="text">
    <img src="forum_images/upload_files.gif" alt="File Uploading Animation" />
    <br /><br />
    <iframe src="AspUpload_ProgressBar.asp?to=10&PID=<% = strPID %>" id="UploadProgress" noresize scrolling="0" frameborder="0" framespacing="0" width="369" height="100"></iframe>
  </td>
 </tr>
</table>
</body>
</html>