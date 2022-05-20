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
<title>Fonts</title>
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
function mouseClick(fontType){
	self.parent.initialiseCommand(fontType);
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
.pStyle { font-size: 14px; color: #000000; }
</style>
</head>
<body bgcolor="#FFFFFF" leftmargin="2" topmargin="2" marginwidth="2" marginheight="2">
<table width="100%"  border="0" cellspacing="0" cellpadding="3">
  <tr onMouseover="overItem(this)" onMouseout="outItem(this)" OnClick="mouseClick('Arial, Helvetica, sans-serif')" style="cursor: default;">
    <td width="100%" class="pStyle" style="font-family: Arial, Helvetica, sans-serif;">Arial</td>
  </tr>
  <tr onMouseover="overItem(this)" onMouseout="outItem(this)" OnClick="mouseClick('Times New Roman, Times, serif')" style="cursor: default;">
    <td class="pStyle" style="font-family: Times New Roman, Times, serif;">Times New Roman</td>
  </tr>
  <tr onMouseover="overItem(this)" onMouseout="outItem(this)" OnClick="mouseClick('Courier New, Courier, mono')" style="cursor: default;">
    <td class="pStyle" style="font-family: Courier New, Courier, mono;">Courier</td>
  </tr>
  <tr onMouseover="overItem(this)" onMouseout="outItem(this)" OnClick="mouseClick('Georgia, Times New Roman, Times, serif')" style="cursor: default;">
    <td class="pStyle" style="font-family: Georgia, Times New Roman, Times, serif;">Georgia</td>
  </tr>
  <tr onMouseover="overItem(this)" onMouseout="outItem(this)" OnClick="mouseClick('Verdana, Arial, Helvetica, sans-serif')" style="cursor: default;">
    <td class="pStyle" style="font-family: Verdana, Arial, Helvetica, sans-serif;">Verdana</td>
  </tr>
  <tr onMouseover="overItem(this)" onMouseout="outItem(this)" OnClick="mouseClick('Geneva, Arial, Helvetica, sans-serif')" style="cursor: default;">
    <td class="pStyle" style="font-family: Geneva, Arial, Helvetica, sans-serif;">Geneva</td>
  </tr>
</table>
</body>
</html>