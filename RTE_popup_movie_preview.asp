<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
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


'Clean up
Call closeDatabase()

'If the page is opening in SSL then keep it running in SSL
If Request.ServerVariables("HTTPS") = "on" Then blnSslEnabledPage = True


Response.AddHeader "pragma","cache"
Response.AddHeader "cache-control","public"
Response.CacheControl = "Public"


Dim strBBcode 			'Holds the bbcode
Dim strMoviePreview		'Holds the HTML formated code
Dim lngStartPos			'Holds search start postions
Dim lngEndPos			'Holds end start postions


'Intiliase the response with a no preview avilable
strMoviePreview = "<br /><br />" & strTxtNoPreviewAvailableForLink


'Read in the BBcode
strBBcode = Request.QueryString("BBcode")


'Convert BBcode into movie
If blnFlashFiles Then
	'Flash
	If InStr(1, strBBcode, "[FLASH", 1) > 0 AND InStr(1, strBBcode, "[/FLASH]", 1) > 0 Then 
		
		'Remove any extra code passed in the querystring
		lngStartPos = InStr(1, strBBcode, "[FLASH", 1)
		lngEndPos = InStr(lngStartPos, strBBcode, "[/FLASH]", 1) + 8
		strBBcode = Trim(Mid(strBBcode, lngStartPos, lngEndPos-lngStartPos))
		
		'Format the Adobe Flash movie preview
		strMoviePreview = formatFlash(strBBcode)
	
			
	'YouTube
	ElseIf InStr(1, strBBcode, "[TUBE]", 1) > 0 AND InStr(1, strBBcode, "[/TUBE]", 1) > 0 Then 
		
		'Get ride of whole link and just leave file name
		strBBcode = Replace(strBBcode, "http://", "https://")
		strBBcode = Replace(strBBcode, "https://youtu.be/", "")
		strBBcode = Replace(strBBcode, "https://www.youtube.com/watch?v=", "")
		strBBcode = Replace(strBBcode, "https://www.youtube.com/v/", "")
		strBBcode = Replace(strBBcode, "&feature=rec-fresh", "")
		
		'Remove any extra code passed in the querystring
		lngStartPos = InStr(1, strBBcode, "[TUBE]", 1)
		lngEndPos = InStr(lngStartPos, strBBcode, "[/TUBE]", 1) + 7
		strBBcode = Trim(Mid(strBBcode, lngStartPos, lngEndPos-lngStartPos))
		
		'Format the YouTube movie preview
		strMoviePreview = formatYouTube(strBBcode)
		
	End If		
End If


%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>No Preview</title>

<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write(vbCrLf & "<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Rich Text Editor " & _
vbCrLf & "Info: http://www.richtexteditor.org" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>

</head>
<style type="text/css">
<!--
.text {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 13px;
	color: #000000;
}
html,body { 
	border: 0px; 
}
-->
</style>
<body bgcolor="#FFFFFF" style="margin:0px;">
<div align="center" class="text"><% = strMoviePreview %></div>
</body>
</html>