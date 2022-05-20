<% @ Language=VBScript %>
<% Option Explicit %>
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



'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write(vbCrLf & "<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Rich Text Editor " & _
vbCrLf & "Info: http://www.richtexteditor.org" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******


'***********************************************************
'*** JavaScript inserting element in Mozilla and Opera *****
'***********************************************************

	%>
//Function to insert element in required position
function insertElementPosition(docWindow, insertElement){

      	var area = docWindow.getSelection();
      	var range = area.getRangeAt(0);
      	var editorPosition = range.startContainer;
      	var pos = range.startOffset;

      	area.removeAllRanges();
      	range.deleteContents();
      	range = document.createRange();

	if (editorPosition.nodeType==3 && insertElement.nodeType==3) {
	        	editorPosition.insertData(pos, insertElement.nodeValue);
	        	try{range.setEnd(editorPosition, pos + insertElement.length);}catch(exception){}
	        	try{range.setStart(editorPosition, pos + insertElement.length);}catch(exception){}

	}else{
	        	var afterElement;
	        	if (editorPosition.nodeType==3){

	          		var textElement = editorPosition;
	          		var text = textElement.nodeValue;
	          		var textBefore = text.substr(0,pos);
	          		var textAfter = text.substr(pos);
	          		var beforeNode = document.createTextNode(textBefore);
	          		var afterElement = document.createTextNode(textAfter);

				editorPosition = textElement.parentNode;
	          		editorPosition.insertBefore(afterElement, textElement);
	          		editorPosition.insertBefore(insertElement, afterElement);
	          		editorPosition.insertBefore(beforeNode, insertElement);
	          		editorPosition.removeChild(textElement);
	        	}else{
	          		afterElement = editorPosition.childNodes[pos];
	          		editorPosition.insertBefore(insertElement, afterElement);
	        	}
	        	try{range.setEnd(afterElement, 0);}catch(exception){}
	        	try{range.setStart(afterElement, 0);}catch(exception){}

      	}
      	area.removeAllRanges();
      	area.addRange(range);
}