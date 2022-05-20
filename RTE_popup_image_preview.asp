<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
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


Response.AddHeader "pragma","cache"
Response.AddHeader "cache-control","public"
Response.CacheControl = "Public"

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>No Preview</title>

<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write(vbCrLf & "<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Rich Text Editor " & _
vbCrLf & "Info: http://www.richtexteditor.org" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>

<script language="JavaScript">

//function to upadte image properties
function imageProperties(oImage){

	if (document.getElementById('prevFile').width != 1 && document.getElementById('prevFile').height != 1){
		window.parent.document.getElementById('width').value = document.getElementById('prevFile').width
		window.parent.document.getElementById('height').value = document.getElementById('prevFile').height
	}
}

</script>
<style type="text/css">
<!--
.text {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #000000;
}
html,body { 
	border: 0px; 
}
-->
</style>
</head>
<body bgcolor="#FFFFFF" leftmargin="2" topmargin="2" marginwidth="2" marginheight="2">
<img src="<% = strImagePath %>clear_pixel.gif" id="prevFile" onError="alert('<% = strTxtErrorLoadingPreview %>')" onLoad="imageProperties(this)"><span class="text">Sample text for previewing image alignment; Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas interdum nulla ut arcu varius ornare. Duis neque nulla, porttitor in, interdum nec, tempor nec, erat. Integer wisi erat, consequat at, sagittis et, tristique lacinia, ipsum. Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas interdum nulla ut arcu varius ornare. Duis neque nulla, porttitor in, interdum nec, tempor nec, erat. Integer wisi erat, consequat at, sagittis et, tristique lacinia, ipsum. Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas interdum nulla ut arcu varius ornare. Duis neque nulla, porttitor in, interdum nec, tempor nec, erat. Integer wisi erat, consequat at, sagittis et, tristique lacinia, ipsum. </span>
</body>
</html>