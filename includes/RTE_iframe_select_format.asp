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

%>
<!-- #include file="browser_page_encoding_inc.asp" -->
<title>Font Format</title>
<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write("<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Rich Text Editor " & _
vbCrLf & "Info: http://www.richtexteditor.org" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>
<script language="JavaScript">
//Function to select font
function mouseClick(fontFormat){
	self.parent.initialiseCommand(fontFormat);
}

//Function to hover item
function overItem(iconItem) {
	iconItem.style.backgroundColor='#CCCCCC';
}

//Function to moving off item
function outItem(iconItem) {
	iconItem.style.backgroundColor='#FFFFFF';
}
</script>
<style type="text/css">
.pStyle { font-family: Arial, Helvetica, sans-serif; color: #000000; font-weight: bold;}
</style>
</head>
<body bgcolor="#FFFFFF" leftmargin="2" topmargin="2" marginwidth="2" marginheight="2">
<table width="100%"  border="0" cellspacing="0" cellpadding="3">
  <tr onMouseover="overItem(this)" onMouseout="outItem(this)" OnClick="mouseClick('<p>')" style="cursor: default;">
    <td width="100%" class="pStyle" style="font-size: 14px; font-weight: normal;">Normal (p)</td>
  </tr>
  <tr onMouseover="overItem(this)" onMouseout="outItem(this)" OnClick="mouseClick('<h1>')" style="cursor: default;">
    <td class="pStyle" style="font-size: 32px;">Heading 1 (H1)</td>
  </tr>
  <tr onMouseover="overItem(this)" onMouseout="outItem(this)" OnClick="mouseClick('<h2>')" style="cursor: default;">
    <td class="pStyle" style="font-size: 24px;">Heading 2 (H2)</td>
  </tr>
  <tr onMouseover="overItem(this)" onMouseout="outItem(this)" OnClick="mouseClick('<h3>')" style="cursor: default;">
    <td class="pStyle" style="font-size: 19px;">Heading 3 (H3)</td>
  </tr>
  <tr onMouseover="overItem(this)" onMouseout="outItem(this)" OnClick="mouseClick('<h4>')" style="cursor: default;">
    <td class="pStyle" style="font-size: 16px;">Heading 4 (H4)</td>
  </tr>
  <tr onMouseover="overItem(this)" onMouseout="outItem(this)" OnClick="mouseClick('<h5>')" style="cursor: default;">
    <td class="pStyle" style="font-size: 13px;">Heading 5 (H5)</td>
  </tr>
  <tr onMouseover="overItem(this)" onMouseout="outItem(this)" OnClick="mouseClick('<h6>')" style="cursor: default;">
    <td class="pStyle" style="font-size: 11px;">Heading 6 (H6)</td>
  </tr>
  <tr onMouseover="overItem(this)" onMouseout="outItem(this)" OnClick="mouseClick('<address>')" style="cursor: default;">
    <td class="pStyle"><address>Address (ADDR)</address></td>
  </tr>
  <tr onMouseover="overItem(this)" onMouseout="outItem(this)" OnClick="mouseClick('<pre>')" style="cursor: default;">
    <td class="pStyle" style="font-weight: normal;"><pre>Formatted (PRE)</pre></td>
  </tr>
</table>
</body>
</html>