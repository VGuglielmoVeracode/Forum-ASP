<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
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




'Clean up
Call closeDatabase()


'If the page is opening in SSL then keep it running in SSL
If Request.ServerVariables("HTTPS") = "on" Then blnSslEnabledPage = True



'Dimension veriables
Dim strLinkType
Dim strHyperlinkType
Dim strHyperlink
Dim strTitle
Dim strWindow
Dim strEmail
Dim strSubject



'If this a post back read in the form elements
If Request.Form("URL") <> "" OR Request.Form("email") <> "" AND Request.Form("postBack") Then
	
	'Get form elements
	strLinkType = Request.Form("selType")
	strHyperlinkType = Request.Form("linkChoice")
	strHyperlink = Request.Form("URL")
	strTitle = Request.Form("Title")
	strWindow = Request.Form("Window")
	strEmail = Request.Form("email")
	strSubject = Request.Form("subject")
	
	'If the http:// part is repeated in the URL then strip it:-
	strHyperlink = Replace(strHyperlink, strHyperlinkType, "", 1, -1, 1)
	
	'Escape characters that will course a crash
	strHyperlink = Replace(strHyperlink, "'", "\'", 1, -1, 1)
	strHyperlink = Replace(strHyperlink, """", "\""", 1, -1, 1)
	strTitle = Replace(strTitle, "'", "\'", 1, -1, 1)
	strTitle = Replace(strTitle, """", "\""", 1, -1, 1)
	strWindow = Replace(strWindow, "'", "\'", 1, -1, 1)
	strWindow = Replace(strWindow, """", "\""", 1, -1, 1)
	strEmail = Replace(strEmail, "'", "\'", 1, -1, 1)
	strEmail = Replace(strEmail, """", "\""", 1, -1, 1)
	strSubject = Replace(strSubject, "'", "\'", 1, -1, 1)
	strSubject = Replace(strSubject, """", "\""", 1, -1, 1)
	
	
	'If this is an email mailto then set the email type to mailto:
	If strLinkType = "email" Then 
		strHyperlinkType = "mailto:"
		strHyperlink = strEmail
		If strSubject <> "" Then strHyperlink = strHyperlink & "?subject=" & strSubject
	End If
End If

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Hyperlink Properties</title>

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
If Request.Form("URL") <> "" OR Request.Form("email") <> "" AND Request.Form("postBack") Then
		
	
	'*********************************************
	'***  	JavaScript for Mozilla & IE	 *****
	'*********************************************
	
	Response.Write(vbCrLf & "editor = window.opener.document.getElementById('WebWizRTE');")
	
	'Mozilla and Opera use different methods than IE to get the selected text (if any)
	If RTEenabled = "Gecko" OR RTEenabled = "opera" Then 
		Response.Write(vbCrLf & vbCrLf & "var selectedRange = editor.contentWindow.window.getSelection();")
	Else	
		Response.Write(vbCrLf & vbCrLf & "var selectedRange = editor.contentWindow.document.selection.createRange();")
	End If	
	


	'If there is a selected area, turn it into a hyperlink
	Response.Write(vbCrLf & vbCrLf & "if (selectedRange != null && selectedRange")
	If RTEenabled = "winIE" Then Response.Write(".text")
	Response.Write(" != ''){")

	'Create hyperlink
	Response.Write(vbCrLf & "	editor.contentWindow.window.document.execCommand('CreateLink', false, '" & strHyperlinkType & strHyperlink & "')")
		
	'Set attributes if required
	If (strLinkType = "link" AND (strTitle <> "" OR strWindow <> "")) OR (strLinkType = "email" AND strSubject <> "") Then
		
		'Set hyperlink attributes
		Response.Write(vbCrLf & vbCrLf & "	var hyperlink = editor.contentWindow.window.document.getElementsByTagName('a');" & _
			       vbCrLf & "	for (var i=0; i < hyperlink.length; i++){" & _
			       vbCrLf & "		if (hyperlink[i].getAttribute('href').search('" & strHyperlinkType & Replace(strHyperlink, "?", "\\?", 1, -1, 1) & "') != -1){")
		
		'Set title, window, subject if required	       
		If strLinkType = "link" AND strTitle <> "" Then Response.Write(vbCrLf & "			hyperlink[i].setAttribute('title','" & strTitle & "');")
		If strLinkType = "link" AND strWindow <> "" Then Response.Write(vbCrLf & "			hyperlink[i].setAttribute('target','" & strWindow & "');")
			       
		Response.Write(vbCrLf & "		}" & _
			       vbCrLf & "	}")
	End If
	
	
	
	'Else no selected area so use the hyperlink text as the displayed text
	Response.Write(vbCrLf & "}else{")
	
	'Tell that we are maiing a hyperlink 'a'
	Response.Write(vbCrLf & vbCrLf & "	hyperlink = editor.contentWindow.document.createElement('a');")
	
	'Create the hyperlink atrtibutes
	Response.Write(vbCrLf & vbCrLf & "	hyperlink.setAttribute('href', '" & strHyperlinkType & strHyperlink & "');")
	If strLinkType = "link" AND strTitle <> "" Then Response.Write(vbCrLf & "	hyperlink.setAttribute('title', '" & strTitle & "');")
	If strLinkType = "link" AND strWindow <> "" Then Response.Write(vbCrLf & "	hyperlink.setAttribute('target', '" & strWindow & "');")
	
	'Use the text eentered for the link to be a child of the a tag so that it is the screen display
	Response.Write(vbCrLf & "	hyperlink.appendChild(editor.contentWindow.document.createTextNode('" & strHyperlinkType & strHyperlink & "'));")
	
	'If this is Mozilla or Opera then we need to call insertElementPosition to find where to place the image
     	If RTEenabled = "Gecko" OR RTEenabled = "opera" Then
		
	Response.Write(vbCrLf & vbCrLf & "	try{" & _
				vbCrLf & "		insertElementPosition(editor.contentWindow, hyperlink);" & _
				vbCrLf & "	}catch(exception){" & _
				vbCrLf & "		alert('" & strTxtErrorInsertingObject & "');" & _
				vbCrLf & "		editor.contentWindow.focus();" & _
				vbCrLf & "	}")
	
	'Else this is IE so placing the link is simpler
	Else
		Response.Write(vbCrLf & vbCrLf & "	try{" & _
					vbCrLf & "		editor.contentWindow.focus();" & _
					vbCrLf & "		editor.contentWindow.document.selection.createRange().pasteHTML(hyperlink.outerHTML);" & _
					vbCrLf & "	}catch(exception){" & _
					vbCrLf & "		alert('" & strTxtErrorInsertingObject & "');" & _
					vbCrLf & "		editor.contentWindow.focus();" & _
					vbCrLf & "	}")
	End If
	
	Response.Write(vbCrLf & "}")
	


	
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

function initialise(){
<%

	'Mozilla and Opera use different methods than IE to get the selected text
	If RTEenabled = "Gecko" OR RTEenabled = "opera" Then
		Response.Write("	var selectedRange = window.opener.document.getElementById('WebWizRTE').contentWindow.window.getSelection().toString();")
	Else
		Response.Write("	var selectedRange = window.opener.document.getElementById('WebWizRTE').contentWindow.document.selection.createRange().text; ")
	End If
%>
	//Use editor selected range to fill text boxes
	if (selectedRange != undefined){
		selectedRange = selectedRange.replace(/http:\/\//i, '');
		selectedRange = selectedRange.replace(/https:\/\//i, '');
		document.getElementById('URL').value = selectedRange;
		document.getElementById('email').value = selectedRange;
	}
	if (document.getElementById('URL').value==''){
		document.getElementById('Submit').disabled=true;
	}
	
	self.focus();
}


//Function to preview URL
function showPreview(linkSelection){
	if (linkSelection.options[linkSelection.selectedIndex].value=="http://" || linkSelection.options[linkSelection.selectedIndex].value=="https://"){
		try {
			document.getElementById("previewLink").contentWindow.location.href =(linkSelection.options[linkSelection.selectedIndex].value + document.getElementById("URL").value);
		}catch(exception){
		}
	
	}else{
		document.getElementById("previewLink").contentWindow.location.href="RTE_popup_link_preview.asp?b=0";
	
	}
}

//Disable preview button for some links
function disablePreview(linkSelection){
	if (linkSelection.options[linkSelection.selectedIndex].value=="http://" || linkSelection.options[linkSelection.selectedIndex].value=="https://"){
		document.getElementById("preview").disabled=false;
		document.getElementById("previewLink").contentWindow.location.href="RTE_popup_link_preview.asp";
		
	}else{
		document.getElementById("preview").disabled=true;
		document.getElementById("previewLink").contentWindow.location.href="RTE_popup_link_preview.asp?b=0";
	}
}

//Function swap link type
function swapLinkType(selType){
	if (selType.value == "email"){
		document.getElementById("hyperlink").style.display="none";
    		document.getElementById("mailLink").style.display="block";<%
    		
'If this is Gekco based browser or Opera the element needs to be set to visable
If RTEenabled = "Gecko" OR RTEenabled = "opera" Then Response.Write(vbCrLf & "		document.getElementById(""mailLink"").style.visibility=""visable"";") 		
    		%>
    		
	}else{
		document.getElementById("mailLink").style.display="none";
		document.getElementById("hyperlink").style.display="block";<%
		
'If this is Gekco based browser or Opera the element needs to be set to visable
If RTEenabled = "Gecko" OR RTEenabled = "opera" Then Response.Write(vbCrLf & "		document.getElementById(""hyperlink"").style.visibility=""visable"";") 		
    		%>
	}
}
</script>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
</head>
<body style="margin:0px;" OnLoad="initialise();">
<table width="100%"  border="0" cellpadding="3" cellspacing="0" class="RTEtableTopRow">
  <form method="post" name="frmLinkInsrt">
    <tr class="RTEtableTopRow">
      <td colspan="2"><h1><% = strTxtHyperlinkProperties %></h1></td>
    </tr>
    <tr>
      <td colspan="2" class="RTEtableRow"><table width="100%" border="0" align="center" cellpadding="1" cellspacing="0">
        <tr>
          <td><table width="100%" border="0" cellpadding="2" cellspacing="0">
            <tr>
              <td width="20%" align="right" class="text"><% = strTxtLinkType %>:</td>
              <td width="80%"><select name="selType" id="selType" onchange="swapLinkType(this)">
                  <option value="link" selected>Hyperlink</option>
                  <option value="email">Email</option>
              </select></td>
            </tr>
          </table></td>
        </tr>
        <tr>
          <td height="240">            
          <span id="hyperlink">
            <table width="100%" border="0" cellpadding="2" cellspacing="0">
              <tr>
                <td width="20%" align="right" class="text"><% = strTxtAddress %>:</td>
                <td width="80%">
                 <select name="linkChoice" id="linkChoice" onchange="disablePreview(this)">
                   <option value="http://" selected>http://</option>
                   <option value="https://">https://</option>
                   <option value="ftp://">ftp://</option>
                   <option value="file://">file://</option>
                   <option value="news://">news://</option>
                   <option value="telnet://">telnet://</option>
                 </select>
                  <input name="URL" type="text" id="URL" size="27" onchange="document.getElementById('Submit').disabled=false;" onkeypress="document.getElementById('Submit').disabled=false;">
                <input name="preview" type="button" id="preview" value="<% = strTxtPreview %>" onclick="showPreview(document.getElementById('linkChoice'))">
                </td>
              </tr>
              <tr>
                <td align="right" valign="top" class="text"><% = strTxtPreview %>:</td>
                <td><iframe src="RTE_popup_link_preview.asp" id="previewLink" width="98%" height="150px" style="border: #A5ACB2 1px solid;background-color: #FFFFFF;"></iframe></td>
              </tr>
          </table>
           </span>
           <span id="mailLink" style="display:none">
            <table width="100%" border="0" cellpadding="2" cellspacing="0">
              <tr>
                <td align="right" class="text"><% = strTxtEmail %>:</td>
                <td><input name="email" type="text" id="email" size="40" onfocus="document.forms.frmLinkInsrt.Submit.disabled=false;"></td>
              </tr>
              <tr>
                <td width="20%" align="right" class="text"><% = strTxtSubject %>:</td>
                <td width="80%"><input name="subject" type="text" id="subject" size="40" maxlength="50"></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>&nbsp;<br /><br /><br /><br /><br /><br /><br /><br /></td>
              </tr>
            </table>
           </span>            </td>
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