<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="RTE_configuration/RTE_setup.asp" -->
<!--#include file="includes/emoticons_inc.asp" -->
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




Response.Buffer = True

'Clean up
Call closeDatabase()


'If emoticons are enabled
If blnEmoticons Then

	'Declare variables
	Dim intIndexPosition		'Holds the idex poistion in the emiticon array
	Dim intNumberOfOuterLoops	'Holds the outer loop number for rows
	Dim intLoop			'Holds the loop index position
	Dim intInnerLoop		'Holds the inner loop number for columns
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="copyright" content="Copyright (C) 2001-2010 Web Wiz" />
<title>Emoticon</title>

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


'If this is Gecko based browser link to JS code for Gecko
If RTEenabled = "Gecko" Then Response.Write(vbCrLf & "<script language=""JavaScript"" src=""RTE_javascript_gecko.asp"" type=""text/javascript""></script>")

%>
		
<script language="JavaScript">

//Function to add the code to the message for the smileys
function AddEmoticon(iconCode) {
 	var txtarea = window.opener.document.frmMessageForm.message;
 	iconCode = ' ' + iconCode + ' ';
 	if (txtarea.createTextRange && txtarea.caretPos) {
  		var caretPos = txtarea.caretPos;
  		caretPos.text = caretPos.text.charAt(caretPos.text.length - 1) == ' ' ? iconCode + ' ' : iconCode;
  		txtarea.focus();
 	} else {
  		txtarea.value  += iconCode;
  		txtarea.focus();
 	}
	window.close();
}

//Function to hover emoticon
function overIcon(iconItem) {
	
	iconItem.style.backgroundColor='#CCCCCC';
	document.getElementById("emotImage").src = iconItem.id;
	document.getElementById("emotName").value = iconItem.title;
}


//Function to moving off emoticon
function outIcon(iconItem) {
	
	iconItem.style.backgroundColor='';
	document.getElementById("emotImage").src = '<% = strImagePath %>blank.gif';
	document.getElementById("emotName").value = '';
}
</script>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
<style type="text/css">
.iconContainer {
	height:280px;
	overflow-x: hidden;
	overflow-y: auto;
}
</style>
</head>
<body bgcolor="#FFFFFF" text="#000000" marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" OnLoad="self.focus();">
<table width="100%"  border="0" cellpadding="3" cellspacing="0" class="tableTopRow">
  <form method="post" name="frmImageInsrt">
    <tr class="tableTopRow">
      <td colspan="2"><h1><% = strTxtEmoticons %></h1></td>
    </tr>
    <tr>
      <td width="80%" class="RTEtableRow">
       <div class="iconContainer">
        <table width="100%" border="0" cellspacing="1" cellpadding="0"><%

	'Intilise the index position (we are starting at 1 instead of position 0 in the array for simpler calculations)
	intIndexPosition = 1
	
	'Calcultae the number of outer loops to do
	intNumberOfOuterLoops = UBound(saryEmoticons) / 6
	
	'If there is a remainder add 1 to the number of loops
	If UBound(saryEmoticons) MOD 6 > 0 Then intNumberOfOuterLoops = intNumberOfOuterLoops + 1
	
	'Loop throgh th list of emoticons
	For intLoop = 1 to intNumberOfOuterLoops
	      
	
		Response.Write(vbCrLf & "	 <tr>")
	
		'Loop throgh th list of emoticons
		For intInnerLoop = 1 to 6  
		
			'If there is nothing to display show an empty box
			If intIndexPosition > UBound(saryEmoticons) Then 
				Response.Write(vbCrLf & "          <td width=""45"" height=""45"" class=""RTEbutton"">&nbsp;</td>") 
	
			'Else show the emoticon
			Else 
				Response.Write(vbCrLf & "          <td width=""45"" height=""45"" class=""RTEbutton"" id=""" & saryEmoticons(intIndexPosition,3) & """ title=""" & saryEmoticons(intIndexPosition,1) & """  onMouseover=""overIcon(this)"" onMouseout=""outIcon(this)"" OnClick=""AddEmoticon('" & saryEmoticons(intIndexPosition,2) & "')"" align=""center"" style=""cursor: default;""><img src=""" & saryEmoticons(intIndexPosition,3) & """ border=""0"" alt=""" & saryEmoticons(intIndexPosition,1) & """></td>")
	              	End If
	              
	              'Minus one form the index position
	              intIndexPosition = intIndexPosition + 1 
		Next    
		        
		Response.Write(vbCrLf & "	 </tr>")
		
	Next             
%>      </div>
       </table>
      </td>
      <td width="20%" align="center" valign="top" class="tableRow"><table width="65" height="45" border="0" cellpadding="0" cellspacing="1">
        <tr>
          <td class="RTEbutton" align="center"><img src="<% = strImagePath %>blank.gif" name="emotImage" id="emotImage"></td>
        </tr>
      </table>
        <input name="emotName" type="text" class="tableRow" style="border: 0px; text-align: center; font-size:10px;" id="emotName" value="" size="15"></td>
    </tr>
    <tr>
    <td class="tableBottomRow">&nbsp;<%

	'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
	If blnAbout Then
		Response.Write("<span class=""text"" style=""font-size:10px"">Rich Text Editing Software by <a href=""http://www.richtexteditor.org"" target=""_blank"" style=""font-size:10px"">Web Wiz Rich Text Editor</a> version " & strRTEversion & "</span>")
	End If 
	'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******      
      
      %></td>
      <td align="right" class="tableBottomRow"><input type="button" name="cancel" value=" <% = strTxtCancel %> " onclick="window.close()">         
      </td></tr>
  </form>
</table>
</body>
</html>
<%

End If

%>