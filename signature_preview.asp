<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="functions/functions_format_post.asp" -->
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


Dim strSignature 		'Holds the Users Message


'Read in the message to be previewed from the cookie set
strSignature = Mid(Request.Form("signature"), 1, 210)



'Call the function to format posts
strSignature = FormatPost(strSignature)

'Call the function to format forum codes
strSignature = FormatForumCodes(strSignature)

'Check the message for malicious HTML code
strSignature = HTMLsafe(strSignature)


'If there is nothing to preview then say so
If strSignature = "" OR IsNull(strSignature) Then
	strSignature = "<br /><br /><div align=""center"">" & strTxtThereIsNothingToPreview & "</div><br /><br />"
'Else fake a post so signature can be view in a real context
Else
	
	'If the signature contains Flash or YouTube
	If blnFlashFiles Then
		If InStr(1, strSignature, "[FLASH", 1) > 0 AND InStr(1, strSignature, "[/FLASH]", 1) > 0 Then strSignature = formatFlash(strSignature)
	End If
		
	'If YouTube
	If blnYouTube Then
		If InStr(1, strSignature, "[TUBE]", 1) > 0 AND InStr(1, strSignature, "[/TUBE]", 1) > 0 Then strSignature = formatYouTube(strSignature)
	End If
End If


'Take the signature down to 255 characters max to prevent database errors
strSignature = Mid(strSignature, 1, 255)

'Clean up
Call closeDatabase()
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title>Signature Preview</title>
<HTTP-EQUIV="PRAGMA" CONTENT="NO-CACHE" /> 

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
</head>
<body bgcolor="#FFFFFF" text="#000000" marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" OnLoad="self.focus();">
<table width="100%"  border="0" cellpadding="3" cellspacing="0" class="tableTopRow">
  <tr class="tableTopRow">
   <td><h1><% = strTxtPreview %></h1></td>
 </tr>
    <tr>
      <td class="tableRow"><br />
        <table width="98%" border="0" cellspacing="0" cellpadding="1" bgcolor="#000000" align="center">
          <tr>
            <td>
              <table width="100%" border="0" cellspacing="0" cellpadding="2" bgcolor="#FFFFFF" height="250" style="table-layout: fixed;">
                <tr>
                  <td class="text" valign="top">
                    <% = strTxtPostedMessage %>
                  </td>
                </tr>
                <tr>
                  <td valign="top" class="msgLineDevider" style="height:150px;">
                   <!-- Start Signature --> 
	            <div class="msgSignature" style="float: left; overflow: auto;">
	             <% = strSignature %>
		    </div>
		   <!-- End Signature -->
                  </td>
                </tr>
            </table></td>
          </tr>
        </table><%
     
     	'If rel=nofollow the display a message
     	If blnNoFollowTagInLinks Then Response.Write("<br /><span class=""smText"">" & strTxtNoFollowAppliedToAllLinks & "</span>")
%> 
      <br />      
     </td>
    </tr>
    <tr>
      <td align="center" class="tableBottomRow">
        <input type="button" name="ok" onclick="javascript:window.close();" value="     <% = strTxtCloseWindow %>     ">
        <br>
        <br>
        <% 
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnLCode Then
	Response.Write("<span class=""text"" style=""font-size:10px"">Forum Software by <a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank""  style=""font-size:10px"">Web Wiz Forums&reg;</a> version " & strVersion & "</span>")
	Response.Write("<br /><span class=""text"" style=""font-size:10px"">Copyright &copy;2001-2019 Web Wiz Ltd.</span>")
End If 
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>
      <br><br>
      </td>
    </tr>
</table>
</body>
</html>