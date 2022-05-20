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




%>
<script language="javascript" src="includes/admin_javascript_v9.js" type="text/javascript"></script>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
</head>
<body>
<table width="100%" border="0" cellspacing="4" cellpadding="0">
 <tr>
  <td colspan="2" valign="top">
   <table align="center" cellpadding="2" cellspacing="1" class="tableBorder" style="height:50px;width:100%">
    <tr class="tableRow">
     <td><span style="float:right;"><a href="admin_logout.asp?XID=<% = getSessionItem("KEY") & strQsSID2 %>"><strong>Control Panel Logout</strong></a></span><img src="<% = strTitleImage %>" border="0" /> <h1>Control Panel<% If blnDemoMode Then  Response.write(" ***** DEMO MODE ****") %></h1></td>
    </tr>
   </table>
  </td>
 </tr>
 <tr valign="top">
  <td width="13%"><!--#include file="admin_index_inc.asp" --></td>
  <td width="87%" align="center">