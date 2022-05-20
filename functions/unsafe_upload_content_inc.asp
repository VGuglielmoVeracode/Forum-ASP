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




'Dimension variables
Dim saryUnSafeHTMLtags(25)	 'If the following contents is found inside an allowed file type that is not text/HTML based then it could be XSS

'Initalise array values
saryUnSafeHTMLtags(0) = "javascript"
saryUnSafeHTMLtags(1) = "vbscript"
saryUnSafeHTMLtags(2) = "jscript"
saryUnSafeHTMLtags(3) = "object"
saryUnSafeHTMLtags(4) = "applet"
saryUnSafeHTMLtags(5) = "embed"
saryUnSafeHTMLtags(6) = "onload"
saryUnSafeHTMLtags(7) = "onclick"
saryUnSafeHTMLtags(8) = "ondblclick"
saryUnSafeHTMLtags(9) = "onkeyup"
saryUnSafeHTMLtags(10) = "onkeydown"
saryUnSafeHTMLtags(11) = "onkeypress"
saryUnSafeHTMLtags(12) = "onmouseenter"
saryUnSafeHTMLtags(13) = "onmouseleave"
saryUnSafeHTMLtags(14) = "onmousemove"
saryUnSafeHTMLtags(15) = "onmouseout"
saryUnSafeHTMLtags(16) = "onmouseover"
saryUnSafeHTMLtags(17) = "onrollover"
saryUnSafeHTMLtags(18) = "onmouse"
saryUnSafeHTMLtags(19) = "onchange"
saryUnSafeHTMLtags(20) = "onunload"
saryUnSafeHTMLtags(21) = "onsubmit"
saryUnSafeHTMLtags(22) = "onselect"
saryUnSafeHTMLtags(23) = "onfocus"
saryUnSafeHTMLtags(24) = "onblur"
saryUnSafeHTMLtags(25) = "onreset"




'If you add more don't forget to increase the number in the Dim statement at the top!
%>