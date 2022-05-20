<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="RTE_configuration/RTE_setup.asp" -->
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



'Dimension veriables
Dim strMovieType
Dim strYouTubeFile
Dim strAdobeFlashURL
Dim intAdobeFlashWidth
Dim intAdobeFlashHeight
Dim strBBcode


strBBcode = ""


'If this a post back read in the form elements
If Request.Form("tubeFile") <> "" OR Request.Form("flashURL") <> "" AND Request.Form("postBack") Then
	
	'Get movie type
	strMovieType = Request.Form("selType")
	
	'If Adobe Flash
	If strMovieType = "Flash" Then 
		
		'Get form elements
		strAdobeFlashURL = Trim(Request.Form("flashURL"))
		If isNumeric(Request.Form("flashWidth")) Then intAdobeFlashWidth = Request.Form("flashWidth") Else intAdobeFlashWidth = 250
		If isNumeric(Request.Form("flashHeight")) Then intAdobeFlashHeight = Request.Form("flashHeight") Else intAdobeFlashHeight = 250
			
		'Create BBcode
		strBBcode = "[FLASH WIDTH=" & intAdobeFlashWidth & " HEIGHT=" & intAdobeFlashHeight & "]" & strAdobeFlashURL & "[/FLASH]"
			
	
	'Else YouTube		
	Else
		'Get form elements
		strYouTubeFile = Trim(Request.Form("tubeFile"))
		
		'Get ride of whole link and just leave file name
		strYouTubeFile = Replace(strYouTubeFile, "http://", "https://")
		strYouTubeFile = Replace(strYouTubeFile, "https://youtu.be/", "")
		strYouTubeFile = Replace(strYouTubeFile, "https://www.youtube.com/watch?v=", "")
		strYouTubeFile = Replace(strYouTubeFile, "https://www.youtube.com/v/", "")
		strYouTubeFile = Replace(strYouTubeFile, "&feature=rec-fresh", "")
		
		
		'Create BBcode
		strBBcode = "[TUBE]" & strYouTubeFile & "[/TUBE]"
	End If
End If

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Movie Properties</title>

<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write(vbCrLf & "<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Rich Text Editor(TM) ver. " & strRTEversion & "" & _
vbCrLf & "Info: http://www.richtexteditor.org" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******


'If this is Gecko or Opera based browser link to JS code for Gecko
If RTEenabled = "Gecko" OR RTEenabled = "opera" Then Response.Write(vbCrLf & "<script language=""JavaScript"" src=""RTE_javascript_gecko.asp"" type=""text/javascript""></script>")
	
%>
<script language="JavaScript">
<%

'If this a post back write javascript
If strBBcode <> "" AND Request.Form("postBack") Then
		
	
	Response.Write(vbCrLf & vbCrLf & "		editor = window.opener.document.getElementById('WebWizRTE');")
		
	'If this is Mozilla or Opera then we need to call insertElementPosition to find where to place the bbcode
	If RTEenabled = "Gecko" OR RTEenabled = "opera" Then 
			
		Response.Write(vbCrLf & vbCrLf & "	try{" & _
					vbCrLf & "		insertElementPosition(editor.contentWindow, editor.contentWindow.document.createTextNode('" & strBBcode & "'));" & _
					vbCrLf & "	}catch(exception){" & _
					vbCrLf & "		alert('" & strTxtErrorInsertingObject & "');" & _
					vbCrLf & "		editor.contentWindow.focus();" & _
					vbCrLf & "	}")
		
	'Else this is IE so placing the bbcode is simpler
	Else
		Response.Write(vbCrLf & vbCrLf & "	try{" & _
					vbCrLf & "		editor.contentWindow.focus();" & _
					vbCrLf & "		editor.contentWindow.document.selection.createRange().pasteHTML('" & strBBcode & "');" & _
					vbCrLf & "	}catch(exception){" & _
					vbCrLf & "		alert('" & strTxtErrorInsertingObject & "');" & _
					vbCrLf & "		editor.contentWindow.focus();" & _
					vbCrLf & "	}")
	End If


	
	'Set focus
	'If Opera change the focus method
	If RTEenabled = "opera" Then
		
		Response.Write(vbCrLf & "	editor.focus();")
	Else
		Response.Write(vbCrLf & "	editor.contentWindow.focus();")
	End If
	
	'Close window
	Response.Write(vbCrLf & "window.close();")	
End If



%>


//Function to preview movie
function showPreview(movieType){
	if (movieType.options[movieType.selectedIndex].value=="Flash") {
		document.getElementById("previewFlash").contentWindow.location.href="RTE_popup_movie_preview.asp?BBcode=[FLASH WIDTH=" + escape(document.getElementById("flashWidth").value) + " HEIGHT=" + document.getElementById("flashHeight").value + "]" + document.getElementById("flashURL").value + "[/FLASH]";
	}else{
		document.getElementById("previewYouTube").contentWindow.location.href="RTE_popup_movie_preview.asp?BBcode=[TUBE]" + escape(document.getElementById("tubeFile").value) + "[/TUBE]";
	
	}
}

//Function swap movie type
function swapLinkType(selType){
	if (selType.value == "Flash"){
		document.getElementById("YouTube").style.display="none";
    		document.getElementById("Flash").style.display="block";<%
    		
'If this is Gekco based browser or Opera the element needs to be set to visable
If RTEenabled = "Gecko" OR RTEenabled = "opera" Then Response.Write(vbCrLf & "		document.getElementById(""Flash"").style.visibility=""visable"";") 		
    		%>
    		
	}else{
		document.getElementById("Flash").style.display="none";
		document.getElementById("YouTube").style.display="block";<%
		
'If this is Gekco based browser or Opera the element needs to be set to visable
If RTEenabled = "Gecko" OR RTEenabled = "opera" Then Response.Write(vbCrLf & "		document.getElementById(""YouTube"").style.visibility=""visable"";") 		
    		%>
	}
}
</script>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
</head>
<body style="margin:0px;">
<table width="100%"  border="0" cellpadding="3" cellspacing="0" class="RTEtableTopRow">
  <form method="post" name="frmLinkInsrt">
    <tr class="RTEtableTopRow">
      <td colspan="2"><h1><% = strTxtMovieProperties %></h1></td>
    </tr>
    <tr>
      <td colspan="2" class="RTEtableRow"><table width="100%" border="0" align="center" cellpadding="1" cellspacing="0">
        <tr>
          <td>
           <table width="100%" border="0" cellpadding="2" cellspacing="0">
            <tr>
              <td width="20%" align="right" class="text"><% = strTxtMovieType %>:</td>
              <td width="80%"><select name="selType" id="selType" onchange="swapLinkType(this)"><%
              	
If blnYouTube Then Response.Write(vbCrLf & "                  <option value=""YouTube"" selected>" & strTxtYouTube & "</option>")
If blnFlashFiles Then Response.Write(vbCrLf & "                  <option value=""Flash"">" & strTxtFlashFilesImages & "</option>")
              	
                  
%>              
              </select></td>
            </tr>
          </table>
         </td>
        </tr>
        <tr>
          <td height="400"><%

If blnYouTube Then

%>           
          <span id="YouTube">
            <table width="100%" border="0" cellpadding="2" cellspacing="0">
              <tr>
                <td width="20%" align="right" class="text"><% = strTxtYouTubeFileName %>:</td>
                <td width="80%">
                  <input name="tubeFile" type="text" id="tubeFile" size="27" onchange="document.getElementById('Submit').disabled=false;" onkeypress="document.getElementById('Submit').disabled=false;">
                <input name="preview1" type="button" id="preview1" value="<% = strTxtPreview %>" onclick="showPreview(document.getElementById('selType'))">
                </td>
              </tr>
              <tr>
                <td align="right" valign="top" class="text"><% = strTxtPreview %>:</td>
                <td><iframe src="RTE_popup_movie_preview.asp" id="previewYouTube" width="98%" height="360px" style="border: #A5ACB2 1px solid;background-color: #FFFFFF;"></iframe></td>
             </tr>
          </table>
           </span><%
End If

If blnFlashFiles Then
%>
           <span id="Flash"<% If blnYouTube Then Response.Write(" style=""display:none""") %>>
            <table width="100%" border="0" cellpadding="2" cellspacing="0">
              <tr>
                <td width="20%" align="right" class="text"><% = strTxtFlashMovieURL %>:</td>
                <td width="80%">
                  <input name="flashURL" type="text" id="flashURL" size="27" value="http://" onchange="document.getElementById('Submit').disabled=false;" onkeypress="document.getElementById('Submit').disabled=false;">
                <input name="preview2" type="button" id="preview2" value="<% = strTxtPreview %>" onclick="showPreview(document.getElementById('selType'))">
                </td>
              </tr>
              <tr>
                <td align="right" valign="top" class="text"> </td>
                <td><% = strTxtHeight %>: <input name="flashHeight" type="text" id="flashHeight" size="2" value="250">&nbsp;&nbsp;&nbsp;&nbsp;<% = strTxtWidth %>: <input name="flashWidth" type="text" id="flashWidth" size="2" value="250"></td>
              </tr>
              <tr>
                <td align="right" valign="top" class="text"><% = strTxtPreview %>:</td>
                <td><iframe src="RTE_popup_movie_preview.asp" id="previewFlash" width="98%" height="325px" style="border: #A5ACB2 1px solid;background-color: #FFFFFF;"></iframe></td>
              </tr>
              </tr>
            </table>
           </span><%
End If

%>            
         </td>
        </tr>
      </table></td>
    </tr>
    <tr>
    <td class="RTEtableBottomRow" valign="top">&nbsp;<%

'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnAbout Then
	Response.Write("<span class=""text"" style=""font-size:10px""><a href=""http://www.richtexteditor.org"" target=""_blank"" style=""font-size:10px"">Web Wiz Rich Text Editor</a> version " & strRTEversion & "</span>")
End If 
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******      
      
      %></td>
      <td align="right" class="RTEtableBottomRow" nowrap valign="top"><input type="hidden" name="postBack" value="true"><input type="submit" id="Submit" name="Submit" value="     <% = strTxtOK %>     ">&nbsp;<input type="button" name="cancel" value=" <% = strTxtCancel %> " onclick="window.close()">
      <br /><br />
     </td>
    </tr>
  </form>
</table>
</body>
</html>