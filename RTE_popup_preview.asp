<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="functions/functions_format_post.asp" -->
<!--#include file="includes/emoticons_inc.asp" -->
<!--#include file="RTE_configuration/RTE_setup.asp" -->
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




'Set the response buffer to true as we maybe redirecting
Response.Buffer = True


'If the page is opening in SSL then keep it running in SSL
If Request.ServerVariables("HTTPS") = "on" Then blnSslEnabledPage = True


Dim strPreviewTextarea 			'Holds the Users Message
Dim strMode


'Make sure this page is not cached
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"

'Read in the message to be previewed
strPreviewTextarea = Mid(Request.Form("pre"), 1, 200000)
strMode = Request.Form("mode")


'Remove start of character put in for firefox bug
strPreviewTextarea = Replace(strPreviewTextarea, chr(2), "")
strPreviewTextarea = Replace(strPreviewTextarea, chr(127), "")


'Place format posts posted with the WYSIWYG Editor (RTE)
If Request.Form("browser") = "RTE" Then 
	
	'Call the function to format WYSIWYG posts
	strPreviewTextarea = WYSIWYGFormatPost(strPreviewTextarea)

'Else standrd editor is used so convert forum codes
Else
	'Call the function to format posts
	strPreviewTextarea = FormatPost(strPreviewTextarea)
End If


'If the user wants forum codes enabled then format the post using them
If Request.Form("forumCodes") Then strPreviewTextarea = FormatForumCodes(strPreviewTextarea)


'Convert text links to hyperlinks
If blnCovertTextLinkstoHyperLink AND blnGroupURLs Then
		
	'If there are no hyperlinks already look to convert text links to hyperlinks
	If InStr(1, strPreviewTextarea, "<a", 1) = 0 AND InStr(1, strPreviewTextarea, "<img", 1) = 0 Then
		strPreviewTextarea = InsertHyperlinks(strPreviewTextarea)
	End If
		
End If
	
	

'Check the message for malicious HTML code
strPreviewTextarea = HTMLsafe(strPreviewTextarea)


'Response.end

strPreviewTextarea = formatInput(strPreviewTextarea)

'If the post contains a Flash

If (blnFlashFiles AND NOT strMode = "PM") OR (strMode = "PM" AND blnPmFlashFiles) Then
	If InStr(1, strPreviewTextarea, "[FLASH", 1) > 0 AND InStr(1, strPreviewTextarea, "[/FLASH]", 1) > 0 Then strPreviewTextarea = formatFlash(strPreviewTextarea)
End If
		
'If YouTube
If (blnYouTube AND NOT strMode = "PM") OR (strMode = "PM" AND blnPmYouTube)  Then
	If InStr(1, strPreviewTextarea, "[TUBE]", 1) > 0 AND InStr(1, strPreviewTextarea, "[/TUBE]", 1) > 0 Then strPreviewTextarea = formatYouTube(strPreviewTextarea)
End If

'Reset server objects
Call closeDatabase()


%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="copyright" content="Copyright (C) 2001-2010 Web Wiz" />
<title>Preview</title>
<HTTP-EQUIV="PRAGMA" CONTENT="NO-CACHE" /> 

<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write(vbCrLf & "<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Rich Text Editor(TM) ver. " & strRTEversion & "" & _
vbCrLf & "Info: http://www.richtexteditor.org" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->" & vbCrLf)
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>

<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />

<%
'If the post contains a quote or code block then format it
'Needs to be called after the skin file include so that colours etc. are carried across
If InStr(1, strPreviewTextarea, "[QUOTE=", 1) > 0 AND InStr(1, strPreviewTextarea, "[/QUOTE]", 1) > 0 Then strPreviewTextarea = formatUserQuote(strPreviewTextarea)
If InStr(1, strPreviewTextarea, "[QUOTE]", 1) > 0 AND InStr(1, strPreviewTextarea, "[/QUOTE]", 1) > 0 Then strPreviewTextarea = formatQuote(strPreviewTextarea)
If InStr(1, strPreviewTextarea, "[CODE]", 1) > 0 AND InStr(1, strPreviewTextarea, "[/CODE]", 1) > 0 Then strPreviewTextarea = formatCode(strPreviewTextarea)
If InStr(1, strPreviewTextarea, "[HIDE]", 1) > 0 AND InStr(1, strPreviewTextarea, "[/HIDE]", 1) > 0 Then strPreviewTextarea = formatHide(strPreviewTextarea)
If InStr(1, strPreviewTextarea, "[SPOILER]", 1) > 0 AND InStr(1, strPreviewTextarea, "[/SPOILER]", 1) > 0 Then strPreviewTextarea = formatSpoiler(strPreviewTextarea, False)

'If there is nothing to preview then say so
If strPreviewTextarea = "" OR IsNull(strPreviewTextarea) OR (InStr(1, strPreviewTextarea, "<br>", 1) = 1 AND Len(strPreviewTextarea) = 8) Then
	strPreviewTextarea = "<br /><br /><div align=""center"">" & strTxtThereIsNothingToPreview & "</div><br /><br />"
End If
%>
</head>
<body  style="margin:0px;" OnLoad="self.focus();">
<table width="100%"  border="0" cellpadding="3" cellspacing="0" class="RTEtableTopRow">
    <tr class="RTEtableTopRow">
      <td><h1><% = strTxtPreview %></h1></td>
    </tr>
    <tr>
      <td class="RTEtableRow"><br />
        <table width="98%" border="0" cellspacing="0" cellpadding="1" align="center">
          <tr>
            <td>
              <table width="100%" border="0" cellspacing="0" cellpadding="2" height="250">
                <tr>
                  <td class="text" valign="top">
                    <!--// /* Message body -->
                    <% = strPreviewTextarea %>
                    <!-- Message body ''"" */ //-->
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
      <td align="center" class="RTEtableBottomRow">
        <input type="button" name="ok" onclick="javascript:window.close();" value="     <% = strTxtCloseWindow %>     ">
        <br>
        <br>
        <% 
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnLCode Then
	Response.Write("<span class=""text"" style=""font-size:10px""><a href=""http://www.richtexteditor.org"" target=""_blank"" style=""font-size:10px"">Web Wiz Rich Text Editor</a> version " & strRTEversion & "</span>")
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