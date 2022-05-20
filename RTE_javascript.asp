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





'Reset Server Objects
Call closeDatabase()

Response.AddHeader "pragma","cache"
Response.AddHeader "cache-control","public"
Response.CacheControl = "Public"

Dim strTextAreaName
Dim strQuickReply

strTextAreaName = Trim(Mid(characterStrip(Request.QueryString("textArea")), 1, 15))
strQuickReply = Trim(Mid(characterStrip(Request.QueryString("QR")), 1, 10))


'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write(vbCrLf & "<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Rich Text Editor(TM) ver. " & strRTEversion & "" & _
vbCrLf & "Info: http://www.richtexteditor.org" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>

var colour;
var htmlOn;
var textAreaName = '<% = strTextAreaName %>';

<%



'***********************************************
'***   JavaScript initialising RTE editor  *****
'***********************************************
%>
//initialise RTE editor
function initialiseWebWizRTE(){

	var textArea = document.getElementById(textAreaName);
	var editor = document.getElementById('WebWizRTE').contentWindow.document;

	function initIframe(){
<%

'IE uses contentEditable instead of designMode to prevent runtime errors in IE6.0.26 to IE6.0.28
'IE uses proprietary attachEvent instead of following the W3C Events module and using addEventListener
'IE SUCKS!!
If RTEenabled = "winIE" Then

%>		
		editor.attachEvent('onkeypress', editorEvents);
		editor.attachEvent('onmousedown', editorEvents);
		document.attachEvent('onmousedown', hideIframes);
		editor.body.contentEditable = true;
<%

'Gekco needs designMode enabled AFTER we listen for events using addEventListener
Else
%>		editor.addEventListener('keypress', editorEvents, true);
		editor.addEventListener('mousedown', editorEvents, true);
		document.addEventListener('mousedown', hideIframes, true);
		editor.designMode = 'on';
<%
End If

%>	}
	setTimeout(initIframe, 300);
	
	//resetting the form
	textArea.form.onreset = function(){
		if (window.confirm('<% = strResetWarningFormConfirm %>')){
	 		editor.body.innerHTML = '';
	 		return true;
	 	}
	return false;
	}
}
<%






'**********************************************
'***   JavaScript create RTE toolbar	  *****
'**********************************************
%>
//Create RTE toolbar
function WebWizRTEtoolbar(formName){
<%
	'Open Iframes

	'Colour Palette iframe
	If blnTextColour OR blnTextBackgroundColour Then
		Response.Write(vbCrLf & "	document.writeln('<iframe width=""260"" height=""165"" id=""colourPalette"" src=""includes/RTE_iframe_colour_palette.asp"" style=""visibility:hidden; position: absolute; left: 0px; top: 0px;"" frameborder=""0"" scrolling=""no""></iframe>');")
	End If

	'Format font Select iframe
	If blnFontStyle AND strQuickReply = "" Then
		Response.Write(vbCrLf & "	document.writeln('<iframe width=""240"" height=""250"" id=""formatFont"" src=""includes/RTE_iframe_select_format.asp"" style=""visibility:hidden; position: absolute; left: 0px; top: 0px; border: 1px solid #000000;"" frameborder=""0"" scrolling=""no""></iframe>');")
	End If

	'Font Select iframe
	If blnFontType  AND strQuickReply = "" Then
		Response.Write(vbCrLf & "	document.writeln('<iframe width=""130"" height=""140"" id=""fontSelect"" src=""includes/RTE_iframe_select_font.asp"" style=""visibility:hidden; position: absolute; left: 0px; top: 0px; border: 1px solid #000000;"" frameborder=""0"" scrolling=""no""></iframe>');")
	End If

	'Font Size iframe
	If blnFontSize  AND strQuickReply = "" Then
		Response.Write(vbCrLf & "	document.writeln('<iframe width=""66"" height=""235"" id=""textSize"" src=""includes/RTE_iframe_select_size.asp"" style=""visibility:hidden; position: absolute; left: 0px; top: 0px; border: 1px solid #000000;"" frameborder=""0"" scrolling=""no""></iframe>');")
	End If

%>
	document.writeln('<table id="toolbar" width="<% If blnMobileBrowser Then Response.Write("100%") Else Response.Write("650") %>" border="0" cellspacing="0" cellpadding="1" class="RTEtoolbar">');
	document.writeln(' <tr>');
	document.writeln('  <td>');<%

'If quick reply load a different toolbar
If strQuickReply = "true" Then
	%><!--#include file="includes/RTE_quick_reply_toolbar_inc.asp" --><%

'If not quick reply, load standard toolbar
Else
	%><!--#include file="includes/RTE_toolbar_inc.asp" --><%
End If
%>
	document.writeln('  </td>');
	document.writeln(' </tr>');
	document.writeln('</table>');
}
<%







'***********************************************
'*** JavaScript for main editor buttons	   *****
'***********************************************
%>
//Function to format text in the text box
function FormatText(command, option){<%


'If this is the Gecko engine then uncomment the following line if you don't wish to use CSS
If RTEenabled = "Gecko" AND blnUseCSS = false Then Response.Write("	document.getElementById('WebWizRTE').contentWindow.document.execCommand(""useCSS"", false, option);")

%>

	var editor = document.getElementById('WebWizRTE');

	//Show iframes
	if ((command == 'forecolor') || (command == 'backcolor') || (command == 'hilitecolor') || (command == 'fontname') || (command == 'formatblock') || (command == 'fontsize')){
		parent.command = command;
		buttonElement = document.getElementById(command);
		switch (command){
			case 'fontname': iframeWin = 'fontSelect'; break;
			case 'formatblock': iframeWin = 'formatFont'; break;
			case 'fontsize': iframeWin = 'textSize'; break;
			default: iframeWin = 'colourPalette';
		}
<%

	'If Opera change the focus method
	If RTEenabled = "opera" Then
		
		Response.Write("		editor.focus();")
	Else
		Response.Write("		editor.contentWindow.focus()")
	End If

%>
		document.getElementById(iframeWin).style.left = getOffsetLeft(buttonElement) + 'px';
		document.getElementById(iframeWin).style.top = (getOffsetTop(buttonElement) + buttonElement.offsetHeight) + 'px';

		if (document.getElementById(iframeWin).style.visibility=='visible'){
			hideIframes();
		}else{
			hideIframes();
			document.getElementById(iframeWin).style.visibility='visible';
		}

		var selectedRange = editor.contentWindow.document.selection;
		if (selectedRange != null){
			range = selectedRange.createRange();
		}
	}<%



'If this is the Gecko or Opera then check the users preferences are set to cut, copy, or paste
If RTEenabled = "Gecko" OR RTEenabled = "opera" Then

%>
	//Paste for AppleWebKit (Safari & Chrome)
	else if ((navigator.userAgent.indexOf('AppleWebKit') > 0) & (command == 'paste')){
	
		alert('<% = strTxtYourBrowserSettingsDoNotPermit %> \'' + command + '\' <% = strTxtPleaseUseKeybordsShortcut %> \(<% = strTxtWindowsUsers %> Ctrl + v, <% = strTxtMacUsers %> Apple + v\)')
	}
	
	//Cut, copy, paste for Gecko
	else if ((command == 'cut') || (command == 'copy') || (command == 'paste')){
		try{
<%

	'If Opera change the focus method
	If RTEenabled = "opera" Then
		
		Response.Write("	  		editor.focus();")
	Else
		Response.Write("	  		editor.contentWindow.focus()")
	End If

%>
	  		editor.contentWindow.document.execCommand(command, false, option);
		}catch(exception){
			switch(command){
				case 'cut': keyboard = 'x'; break;
				case 'copy': keyboard = 'c'; break;
				case 'paste': keyboard = 'v'; break;
			}
			alert('<% = strTxtYourBrowserSettingsDoNotPermit %> \'' + command + '\' <% = strTxtPleaseUseKeybordsShortcut %> \(<% = strTxtWindowsUsers %> Ctrl + ' + keyboard + ', <% = strTxtMacUsers %> Apple + ' + keyboard + '\)')
		}

	}<%
End If



'If the advanced hyperlink is not enabled then display a basic hyperlink function
If blnAdvAdddHyperlink = false Then

%>
	else if (command == 'createLink'){

<%

	'Mozilla and Opera use different methods than IE to get the selected text
	If RTEenabled = "Gecko" OR RTEenabled = "opera" Then
		Response.Write("		var selectedRange = editor.contentWindow.window.getSelection().toString();")
	Else
		Response.Write("		var selectedRange = editor.contentWindow.document.selection.createRange().text; ")
	End If
%>
		if (selectedRange != null && selectedRange != ''){
			//place http infront if not already in selected range
			if (selectedRange.substring(0,4) != 'http'){
				selectedRange = 'http://' + selectedRange
			}

			insertLink = prompt('<% = strTxtEnterHeperlinkURL %>', selectedRange);

			if ((insertLink != null) && (insertLink != '')){
<%

	'If Opera change the focus method
	If RTEenabled = "opera" Then
		
		Response.Write("				editor.focus();")
	Else
		Response.Write("				editor.contentWindow.focus()")
	End If

%>
				editor.contentWindow.document.execCommand('CreateLink', false, insertLink);
			}
		}else{
			alert('<% = strTxtSelectTextToTurnIntoHyperlink %>')
		}
	}<%

End If


'Else none of the other command need extra code so run the command as a bsic execCommand in the editor
%>
	else{
	  	editor.contentWindow.focus();
	  	editor.contentWindow.document.execCommand(command, false, option);
	}
<%

	'If Opera change the focus method
	If RTEenabled = "opera" Then
		
		Response.Write("	  	editor.focus();")
	Else
		Response.Write("	  	editor.contentWindow.focus()")
	End If

%>
}
<%

'***********************************************************************************






'************************************************************************
'*** 	JavaScript for initialise commands (iframe colours etc.)    *****
'************************************************************************
%>
//Function to initialise commands
function initialiseCommand(selection){
	var editor = document.getElementById('WebWizRTE')
<%
'If this is IE then use the following
If RTEenabled = "winIE" Then

	%>
	//retrieve selected range
	var selectedRange = editor.contentWindow.document.selection;
	if (selectedRange!=null){
		selectedRange = selectedRange.createRange();
		selectedRange = range;
		selectedRange.select();
	}<%
End If
%>
	editor.contentWindow.document.execCommand(parent.command, false, selection);
<%

'If Opera change the focus method
If RTEenabled = "opera" Then
	
	Response.Write("	editor.focus();")
Else
	Response.Write("	editor.contentWindow.focus();")
End If

%>
	hideIframes();
}
<%





'*****************************************************
'*** 	JavaScript for switching to HTML view    *****
'*****************************************************
If blnHTMLView Then
%>
//Function to switch to HTML view
function HTMLview(){
	var editor = document.getElementById('WebWizRTE');
	<%

	'If this is IE then use the following
	If RTEenabled = "winIE" Then

	%>
	//WYSIWYG view
	if (htmlOn == true){
		var html = editor.contentWindow.document.body.innerText;
		editor.contentWindow.document.body.innerHTML = html;
		document.getElementById('ToolBar1').style.visibility='visible';
		document.getElementById('ToolBar2').style.visibility='visible';
		htmlOn = false;

	//HTML view
	}else{
		var html = editor.contentWindow.document.body.innerHTML;
		editor.contentWindow.document.body.innerText = html;
    		document.getElementById('ToolBar1').style.visibility='hidden';
    		document.getElementById('ToolBar2').style.visibility='hidden';
    		htmlOn = true;
    	}<%


	'Else for Midas
	Else

	%>
	//WYSIWYG view
	if (htmlOn == true){
		var html = editor.contentWindow.document.body.ownerDocument.createRange();
		html.selectNodeContents(editor.contentWindow.document.body);
		editor.contentWindow.document.body.innerHTML = html.toString();
		document.getElementById('ToolBar1').style.visibility='visible';
		document.getElementById('ToolBar2').style.visibility='visible';
		htmlOn = false;

	//HTML view
	}else{
		var html = document.createTextNode(editor.contentWindow.document.body.innerHTML);
    		editor.contentWindow.document.body.innerHTML = '';
    		editor.contentWindow.document.body.appendChild(html);
    		document.getElementById('ToolBar1').style.visibility='hidden';
    		document.getElementById('ToolBar2').style.visibility='hidden';
    		htmlOn = true;
    	}
<%

	End If


	'If Opera change the focus method
	If RTEenabled = "opera" Then
		
		Response.Write("    		editor.focus();")
	Else
		Response.Write("    		editor.contentWindow.focus()")
	End If
%>
}
<%
End If





'***********************************************
'*** 	JavaScript for print content    *****
'***********************************************
If blnPrint Then
%>
//Function to print editor content
function printEditor(){
<%
	'If this is Gekco or Opera then print method is different to IE
	If RTEenabled = "Gecko" OR RTEenabled = "opera" Then
		Response.Write("	document.getElementById('WebWizRTE').contentWindow.print();")
	Else
		Response.Write("	document.getElementById('WebWizRTE').contentWindow.document.execCommand('Print');")
	End If
%>
}
<%
End If





'***********************************************
'*** 	JavaScript for clear button       *****
'***********************************************
'If new doc is enabled then include the following function
If blnNew Then
%>
//Function to clear editor content
function clearWebWizRTE(){
	 if (window.confirm('<% = strResetWarningEditorConfirm %>')){
	 	document.getElementById('WebWizRTE').contentWindow.document.body.innerHTML='';
<%
	'If Opera change the focus method
	If RTEenabled = "opera" Then
		
		Response.Write("	 	document.getElementById('WebWizRTE').focus();")
	Else
		Response.Write("	 	document.getElementById('WebWizRTE').contentWindow.focus()")
	End If
%>
	 }
}
<%
End If




'***********************************************
'*** 	JavaScript for iframes position    *****
'***********************************************
%>
//Iframe top offset
function getOffsetTop(elm){
	var mOffsetTop = elm.offsetTop;
	var mOffsetParent = elm.offsetParent;
	while(mOffsetParent){
		mOffsetTop += mOffsetParent.offsetTop;
		mOffsetParent = mOffsetParent.offsetParent;
	}
	return mOffsetTop;
}

//Iframe left offset
function getOffsetLeft(elm){
	var mOffsetLeft = elm.offsetLeft;
	var mOffsetParent = elm.offsetParent;
	while(mOffsetParent){
		mOffsetLeft += mOffsetParent.offsetLeft;
		mOffsetParent = mOffsetParent.offsetParent;
	}
	return mOffsetLeft;
}
<%




'***********************************************
'*** 	JavaScript for iframe hidding 	   *****
'***********************************************

Response.Write("//Function to hide iframes" & _
      vbCrLf & "function hideIframes(){")

If blnTextColour OR blnTextBackgroundColour OR blnFontStyle OR blnFontType OR blnFontSize Then

	'Colour Palette iframe
	If blnTextColour OR blnTextBackgroundColour Then

		Response.Write(vbCrLf & "	if (document.getElementById('colourPalette').style.visibility=='visible'){document.getElementById('colourPalette').style.visibility='hidden';}")
	End If

	'Format font Select iframe
	If blnFontStyle AND strQuickReply = "" Then
		Response.Write(vbCrLf & "	if (document.getElementById('formatFont').style.visibility=='visible'){document.getElementById('formatFont').style.visibility='hidden';}")
	End If

	'Font Select iframe
	If blnFontType AND strQuickReply = "" Then
		Response.Write(vbCrLf & "	if (document.getElementById('fontSelect').style.visibility=='visible'){document.getElementById('fontSelect').style.visibility='hidden';}")
	End If

	'Font Size iframe
	If blnFontSize AND strQuickReply = "" Then
		Response.Write(vbCrLf & "	if (document.getElementById('textSize').style.visibility=='visible'){document.getElementById('textSize').style.visibility='hidden';}")
	End If
End If

Response.Write(vbCrLf & "}")




'***********************************************
'***  JavaScript for spell check detection *****
'***********************************************

'If spell checking is enabled load the following function
If blnSpellCheck Then
%>

//Function to perform spell check
function checkspell(){<%

'If IE
If RTEenabled = "winIE" Then

	%>
	try{
		var tmpis = new ActiveXObject('ieSpell.ieSpellExtension');
		tmpis.CheckAllLinkedDocuments(document);
	}
	catch(exception){
		if(exception.number==-2146827859){
			if (confirm('<% = strTxtIeSpellNotDetected %>')){
				window.open('http://www.iespell.com/download.php','DownLoad', '');
			}
		}
		else
			alert('Error Loading ieSpell: Exception ' + exception.number);
	}<%

'Else if this is Gecko browser load different JS
ElseIf RTEenabled = "Gecko" Then

	%>
	if (confirm('<% = strTxtSpellBoundNotDetected %>')){
		window.open('http://spellbound.sourceforge.net/install','DownLoad', '');
	}<%

End If
%>
}<%

End If




'***********************************************
'***   	   Editor Keybord/Mouse Events 	   *****
'***********************************************
%>
//Run Editor Events
function editorEvents(evt){
	var keyCode = evt.keyCode ? evt.keyCode : evt.charCode;
	var keyCodeChar = String.fromCharCode(keyCode).toLowerCase();
<%

'Put in some keybord shortcuts Gecko doesn't have when in RTE mode (IE already has these built in)
If RTEenabled = "Gecko" Then

%>
  	//Keyboard shortcuts
  	if (evt.type=='keypress' && evt.ctrlKey){
  		var kbShortcut;
  		switch (keyCodeChar){
			case 'b': kbShortcut = 'bold'; break;
			case 'i': kbShortcut = 'italic'; break;
			case 'u': kbShortcut = 'underline'; break;
			case 's': kbShortcut = 'strikethrough'; break;
			case 'i': kbShortcut = 'italic'; break;
		}
		if (kbShortcut){
			FormatText(kbShortcut, '');
			evt.preventDefault();
			evt.stopPropagation();
		}
	}
<%
End If


'Prevent double line spacing in IE (IE SUCKS!!!)
'If this is IE then detect if ENTER key is prssed then replace <p> with <div>

'I would replace <p> with <br>, but this then courses problems within tables, ordered lists, and 
'other elements so <div> is used as it creates the same one line effect but without the problems
If RTEenabled = "winIE" AND blnNoIEdblLine Then

%>
	//run if enter key is pressed
	if (evt.type=='keypress' && keyCode==13){
		var editor = document.getElementById('WebWizRTE');
		var selectedRange = editor.contentWindow.document.selection.createRange();
		var parentElement = selectedRange.parentElement();
		var tagName = parentElement.tagName;

		while((/^(a|abbr|acronym|b|bdo|big|cite|code|dfn|em|font|i|kbd|label|q|s|samp|select|small|span|strike|strong|sub|sup|textarea|tt|u|var)$/i.test(tagName)) && (tagName!='HTML')){
			parentElement = parentElement.parentElement;
			tagName = parentElement.tagName;
		}

		//Insert <div> instead of <p>
		if (parentElement.tagName == 'P'||parentElement.tagName=='BODY'||parentElement.tagName=='HTML'||parentElement.tagName=='TD'||parentElement.tagName=='THEAD'||parentElement.tagName=='TFOOT'){
			selectedRange.pasteHTML('<div>');
			selectedRange.select();
			return false;
		}
	}<%
End If

%>
	hideIframes();
	return true;
}

<%
'If emoticons are enabled
If blnEmoticons Then
	
	Response.Write(vbCrLf & "//Function to add emoticon")
	Response.Write(vbCrLf & "function AddEmoticon(iconItem){")
		
	Response.Write(vbCrLf & vbCrLf & "	editor = document.getElementById('WebWizRTE');")
			
	'Tell that we are an image
	Response.Write(vbCrLf & vbCrLf & "	img = editor.contentWindow.document.createElement('img');")
			
	'Set image attributes
	If  blnUseFullURLpath = true Then
		Response.Write(vbCrLf & vbCrLf & "	img.setAttribute('src', '" & strFullURLpathToRTEfiles & "' + iconItem.id);")
	Else
		Response.Write(vbCrLf & vbCrLf & "	img.setAttribute('src', iconItem.id);")
	End If
	Response.Write(vbCrLf & "	img.setAttribute('border', '0');")
	Response.Write(vbCrLf & "	img.setAttribute('alt', iconItem.title);")
	Response.Write(vbCrLf & "	img.setAttribute('align', 'absmiddle');")
			 
	      
	'If this is Mozilla or Opera then we need to call insertElementPosition to find where to place the image
     	If RTEenabled = "Gecko" OR RTEenabled = "opera" Then 
     		
     		Response.Write(vbCrLf & vbCrLf & "	try{" & _
					vbCrLf & "		insertElementPosition(editor.contentWindow, img);" & _
					vbCrLf & "	}catch(exception){" & _
					vbCrLf & "		alert('" & strTxtErrorInsertingObject & "');" & _
					vbCrLf & "		editor.contentWindow.focus();" & _
					vbCrLf & "	}")
			
	'Else this is IE so placing the image is simpler
	Else
		Response.Write(vbCrLf & vbCrLf & "	try{" & _
					vbCrLf & "		editor.contentWindow.focus();" & _
					vbCrLf & "		editor.contentWindow.document.selection.createRange().pasteHTML(img.outerHTML);" & _
					vbCrLf & "	}catch(exception){" & _
					vbCrLf & "		alert('" & strTxtErrorInsertingObject & "');" & _
					vbCrLf & "		editor.contentWindow.focus();" & _
					vbCrLf & "	}")
	End If
				
	'set focus
	'If Opera change the focus method
	If RTEenabled = "opera" Then
		
		Response.Write(vbCrLf & "	editor.focus();")
	Else
		Response.Write(vbCrLf & "	editor.contentWindow.focus();")
	End If
	Response.Write(vbCrLf & "}")
End If
%>