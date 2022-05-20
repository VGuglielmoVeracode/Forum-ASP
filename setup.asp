<% @ Language=VBScript %>
<% Option Explicit %>
<!-- #include file="database/database_connection.asp" -->
<!-- #include file="includes/version_inc.asp" -->
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




'Set the response buffer to true as we maybe redirecting and setting a cookie
Response.Buffer = True


'If the database is already setup then move 'em
If strDatabaseType <> "" Then Response.Redirect("default.asp")

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en" dir="ltr">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Web Wiz Forums Installation Wizard</title>
<meta name="generator" content="Web Wiz Forums" />
<%

Response.Write(vbCrLf  & vbCrLf & "<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Forums(TM) " & _
vbCrLf & "Info: http://www.webwizforums.com" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->" & vbCrLf)
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******

%>
<link href="css_styles/default/default_style.css" rel="stylesheet" type="text/css" />
</head>
<body>
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
  <tr>
    <td><a href="http://www.webwizforums.com"><img src="forum_images/web_wiz_forums.png" border="0" alt="Web Wiz Forums Homepage" title="Web Wiz Forums Homepage" /><br />
      <span class="smText"><img src="forum_images/webwizforums_box.jpg" alt="Web Wiz Forums" width="128" height="140" align="right" /></span><br />
      </a>
      <h1>Web Wiz Forums Installation Wizard - Version
        <% = strVersion %>
      </h1>
      <span class="smText"><br />
      Welcome to the Web Wiz Forums Installation Wizard. This wizard will guide you through the both a installation and upgrade of your Web Wiz Forums Application.<br />
      <br />
      You may navigate through the Wizard using the Next and Previous buttons. On some pages you will see a third button &quot;Test ...&quot;. This button will allow you to test the configuration before you continue, to see the effects of changes.<br />
      <br />
      The first step is to select the type of installation that you require.</span> </td>
  </tr>
</table>
<form id="frmSetup" name="frmSetup" method="post" action="setup_select_database.asp">
  <br />
  <table cellspacing="1" cellpadding="3" class="tableBorder" align="center">
    <tr class="tableLedger">
      <td>Select Installation Type</td>
    </tr>
    <tr>
      <td align="left"  class="tableRow"><label>
        <input type="radio" name="install" id="install1" value="new" checked="checked" />
        New Installation</label></td>
    </tr>
     <tr>
      <td align="left"  class="tableRow"><label>
        <input type="radio" name="install" id="install2" value="upgrade12" />
        Upgrade from Web Wiz Forums 12.x</label></td>
    </tr>
     <tr>
      <td align="left"  class="tableRow"><label>
        <input type="radio" name="install" id="install3" value="upgrade11" />
        Upgrade from Web Wiz Forums 11.x</label></td>
    </tr>
    <tr>
      <td align="left"  class="tableRow"><label>
        <input type="radio" name="install" id="install4" value="upgrade10" />
        Upgrade from Web Wiz Forums 10.x</label></td>
    </tr>
    <tr>
      <td align="left"  class="tableRow"><label>
        <input type="radio" name="install" id="install5" value="upgrade9" />
        Upgrade from Web Wiz Forums 9.x</label></td>
    </tr>
    <tr>
      <td align="left"  class="tableRow"><label>
        <input type="radio" name="install" id="install6" value="upgrade8" />
        Upgrade from Web Wiz Forums 8.x</label></td>
    </tr>
    <tr>
      <td align="left"  class="tableRow"><label>
        <input type="radio" name="install" id="install7" value="upgrade7" />
        Upgrade from Web Wiz Forums 7.x</label></td>
    </tr>
    <tr>
      <td class="tableBottomRow"><table width="100%" border="0" cellspacing="2" cellpadding="0">
          <tr>
            <td width="24%">&nbsp;</td>
            <td width="76%"><input type="submit" name="button" id="button" value="Next &gt;&gt;" /></td>
          </tr>
        </table>
       </td>
    </tr>
  </table>
  <br />
  <div align="center">
    <%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write("<span class=""text"" style=""font-size:10px"">Forum Software by <a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank""  style=""font-size:10px"">Web Wiz Forums&reg;</a> version " & strVersion & "</span>")
Response.Write("<br /><span class=""text"" style=""font-size:10px"">Copyright &copy;2001-2019 Web Wiz Ltd.</span>")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>
  </div>
  <br />
</form>
</body>
</html>
