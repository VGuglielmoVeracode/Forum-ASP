<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="functions/functions_common.asp" -->
<!--#include file="language_files/language_file_inc.asp" -->
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
Response.CacheControl = "Public"

%>
var colour;
var htmlOn;
<%


'*********************************************
'***  	JavaScript for Windows IE5	 *****
'*********************************************


'If this is windows IE 5.0 use different JavaScript functions
If RTEenabled = "winIE5" Then

	
%>
//Function to format text in the text box
function FormatText(command, option){

	//Colour pallete
	if ((command == "forecolor") || (command == "hilitecolor")) {
		
		parent.command = command;
		buttonElement = document.all(command);
		frames.message.focus()
		document.all.colourPalette.style.left = getOffsetLeft(buttonElement) + "px";
		document.all.colourPalette.style.top = (getOffsetTop(buttonElement) + buttonElement.offsetHeight) + "px";
		
		if (document.all.colourPalette.style.visibility=="hidden")
			document.all.colourPalette.style.visibility="visible";
		else {
			document.all.colourPalette.style.visibility="hidden";
		}
		
		//get current selected range
		var sel = frames.message.document.selection; 
		if (sel != null) {
			colour = sel.createRange();
		}
	}

	//Text Format
	frames.message.focus();
  	frames.message.document.execCommand(command, false, option);
  	frames.message.focus();
}

//Function to add image
function AddImage(){	
	imagePath = prompt("<% = strTxtEnterImageURL %>", "http://");				
	
	if ((imagePath != null) && (imagePath != "")){	
		frames.message.focus(); 				
		frames.message.document.execCommand("InsertImage", false, imagePath);
	}
	frames.message.focus();			
}

//Function to switch to HTML view
function HTMLview() {

	//WYSIWYG view
	if (htmlOn == true) {
		var html = frames.message.document.body.innerText;
		frames.message.document.body.innerHTML = html;
		ToolBar1.style.visibility="visible";
		ToolBar2.style.visibility="visible";
		htmlOn = false;
	
	//HTML view
	} else {
		
		var html = frames.message.document.body.innerHTML;
		frames.message.document.body.innerText = html;
    		ToolBar1.style.visibility="hidden";
    		ToolBar2.style.visibility="hidden";
    		htmlOn = true;
    	}	
}

//Function to set colour
function setColor(color) {

	//retrieve selected range
	var sel = frames.message.document.selection; 
	if (sel!=null) {
		var newColour = sel.createRange();
		newColour = colour;
		newColour.select();
	}
		
	frames.message.focus();
	frames.message.document.execCommand(parent.command, false, color);
	frames.message.focus();
	document.all.colourPalette.style.visibility="hidden";
}

//Function to clear form
function ResetForm(){

	if (window.confirm("<% = strResetFormConfirm %>")){
		frames.message.focus();
	 	frames.message.document.body.innerHTML = ""; 
	 	return true;
	 } 
	 return false;		
}

//Function to add smiley
function AddSmileyIcon(imagePath){	
	frames.message.focus();								
	frames.message.document.execCommand("InsertImage", false, imagePath);
}<%






'***********************************************
'*** JavaScript for Win IE5.5+ and Mozilla *****
'***********************************************

'Else use cross browsers RTE JS for all other RTE enabled browsers
Else


%>
//Function to format text in the text box
function FormatText(command, option) {<% 

	'If this is the Gecko engine then uncomment the following line if you don't wish to use CSS
	'If RTEenabled = "Gecko" Then Response.Write("	document.getElementById(""message"").contentWindow.document.execCommand(""useCSS"", false, option);") 

%>	
	//Colour pallete
	if ((command == "forecolor") || (command == "backcolor")) {
		
		parent.command = command;
		buttonElement = document.getElementById(command);
		document.getElementById("message").contentWindow.focus()
		document.getElementById("colourPalette").style.left = getOffsetLeft(buttonElement) + "px";
		document.getElementById("colourPalette").style.top = (getOffsetTop(buttonElement) + buttonElement.offsetHeight) + "px";
		
		if (document.getElementById("colourPalette").style.visibility=="hidden")
			document.getElementById("colourPalette").style.visibility="visible";
		else {
			document.getElementById("colourPalette").style.visibility="hidden";
		}
		
		//get current selected range
		var sel = document.getElementById("message").contentWindow.document.selection; 
		if (sel != null) {
			colour = sel.createRange();
		}
	}<%
	
	 
	'If this is the Gecko then url links are cerated differently
	If RTEenabled = "Gecko" Then	
	
	%>
	//URL link for Gecko
	else if (command == "createLink") {	
		insertLink = prompt("<% = strTxtEnterHeperlinkURL %>", "http://");			
		if ((insertLink != null) && (insertLink != "")) {
			document.getElementById("message").contentWindow.focus()
			document.getElementById("message").contentWindow.document.execCommand("CreateLink", false, insertLink);
			document.getElementById("message").contentWindow.focus()
		}	
	}<%
	End If
 
%>	
	//Text Format
	else {
		document.getElementById("message").contentWindow.focus();
	  	document.getElementById("message").contentWindow.document.execCommand(command, false, option);
		document.getElementById("message").contentWindow.focus();
	}
}

//Function to set colour
function setColor(color) {<%

	'If this is IE then use the following
	If RTEenabled = "winIE" Then
	
	%>
	//retrieve selected range
	var sel = document.getElementById("message").contentWindow.document.selection; 
	if (sel!=null) {
		var newColour = sel.createRange();
		newColour = colour;
		newColour.select();
	}<%
	End If
%>	
	document.getElementById("message").contentWindow.focus();
	document.getElementById("message").contentWindow.document.execCommand(parent.command, false, color);
	document.getElementById("message").contentWindow.focus();
	document.getElementById("colourPalette").style.visibility="hidden";
}


//Function to add image
function AddImage() {
	imagePath = prompt("<% = strTxtEnterImageURL %>", "http://");			
	if ((imagePath != null) && (imagePath != "")) {
		document.getElementById("message").contentWindow.focus()
		document.getElementById("message").contentWindow.document.execCommand("InsertImage", false, imagePath);
	}
	document.getElementById("message").contentWindow.focus()
}

//Function to switch to HTML view
function HTMLview() {
	<%

	'If this is IE then use the following
	If RTEenabled = "winIE" Then
	
	%>
	//WYSIWYG view
	if (htmlOn == true) {
		var html = document.getElementById("message").contentWindow.document.body.innerText;
		document.getElementById("message").contentWindow.document.body.innerHTML = html;
		document.getElementById("ToolBar1").style.visibility="visible";
		document.getElementById("ToolBar2").style.visibility="visible";
		htmlOn = false;
	
	//HTML view
	} else {
		
		var html = document.getElementById("message").contentWindow.document.body.innerHTML;
		document.getElementById("message").contentWindow.document.body.innerText = html;
    		document.getElementById("ToolBar1").style.visibility="hidden";
    		document.getElementById("ToolBar2").style.visibility="hidden";
    		htmlOn = true;
    	}<%
    		
    		
	'Else for Midas (Geckos RTE API)
	Else 
	
	%>
	//WYSIWYG view
	if (htmlOn == true) {
		var html = document.getElementById("message").contentWindow.document.body.ownerDocument.createRange();
		html.selectNodeContents(document.getElementById("message").contentWindow.document.body);
		document.getElementById("message").contentWindow.document.body.innerHTML = html.toString();
		document.getElementById("ToolBar1").style.visibility="visible";
		document.getElementById("ToolBar2").style.visibility="visible";
		htmlOn = false;
	
	//HTML view
	} else {
		var html = document.createTextNode(document.getElementById("message").contentWindow.document.body.innerHTML);
    		document.getElementById("message").contentWindow.document.body.innerHTML = "";
    		document.getElementById("message").contentWindow.document.body.appendChild(html);
    		document.getElementById("ToolBar1").style.visibility="hidden";
    		document.getElementById("ToolBar2").style.visibility="hidden";
    		htmlOn = true;	
    	}<%
    	
    	End If
    	%>		
}

//Function to clear form
function ResetForm() {
	if (window.confirm("<%=strResetFormConfirm%>")) {
		document.getElementById("message").contentWindow.focus()
	 	document.getElementById("message").contentWindow.document.body.innerHTML = ""; 
	 	return true;
	 } 
	 return false;		
}

//Function to add smiley
function AddSmileyIcon(imagePath){	
	document.getElementById("message").contentWindow.focus();							
	document.getElementById("message").contentWindow.document.execCommand("InsertImage", false, imagePath);
}<%

End If



'***********************************************
'*** 	JavaScript for colour palette 	   *****
'***********************************************

%>

//Colour pallete top offset
function getOffsetTop(elm) {
	var mOffsetTop = elm.offsetTop;
	var mOffsetParent = elm.offsetParent;
	while(mOffsetParent){
		mOffsetTop += mOffsetParent.offsetTop;
		mOffsetParent = mOffsetParent.offsetParent;
	}
	return mOffsetTop;
}

//Colour pallete left offset
function getOffsetLeft(elm) {
	var mOffsetLeft = elm.offsetLeft;
	var mOffsetParent = elm.offsetParent;
	while(mOffsetParent) {
		mOffsetLeft += mOffsetParent.offsetLeft;
		mOffsetParent = mOffsetParent.offsetParent;
	}
	return mOffsetLeft;
}

//Function to hide colour pallete
function hideColourPallete() {<%

'If this is win IE 5 use document.all
If RTEenabled = "winIE5" Then 
%>
	document.all.colourPalette.style.visibility="hidden";<%

'For all other browsers use document.getElementById
Else
%>
	document.getElementById("colourPalette").style.visibility="hidden";<%
	
End If
%>
}<%



'***********************************************
'***    	JavaScript for ieSpell 	   *****
'***********************************************

'If this is IE then write the following spel check function
If RTEenabled = "winIE" OR RTEenabled = "winIE5" Then
	
	%>
//Function to perform spell check
function checkspell() {
	try {
		var tmpis = new ActiveXObject("ieSpell.ieSpellExtension");
		tmpis.CheckAllLinkedDocuments(document);
	}
	catch(exception) {
		if(exception.number==-2146827859) {
			if (confirm("<% = strTxtIeSpellNotDetected %>"))
				openWin("http://www.iespell.com/download.php","DownLoad", "");
		}
		else
			alert("Error Loading ieSpell: Exception " + exception.number);
	}
}<%

End If

%>