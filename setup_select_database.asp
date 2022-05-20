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
<title>Web Wiz Forums Select Database</title>
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
      <br />
      </a>
      <h1>Web Wiz Forums Select Database</h1>
      <span class="smText"><br />
      Web Wiz Forums has four different database options, please read each one carefully as the type of database that you choose is important as you can <strong>not</strong> change the type of database you are using once you begin using your forum.<br />
      <br />
      If you are using a 'Hosting Account' then check with your hosting provider which database type your hosting account supports.<br />
      <br />
      Select from the list below which database type you wish to use. </span></td>
  </tr>
</table>
<form id="frmSetup" name="frmSetup" method="post" action="setup_database_connection.asp">
  <br />
  <table cellspacing="1" cellpadding="3" class="tableBorder" align="center">
    <tr class="tableLedger">
      <td colspan="3">Select Database</td>
    </tr>
    <tr>
      <td width="81" align="center"  class="tableRow"><img src="forum_images/MSSQL.gif" alt="Microsoft SQL Server 2005/2008/2012" width="64" height="64" /></td>
      <td colspan="2"  class="tableRow"><label>
        <input name="DbType" type="radio" id="DbType1" value="SQLServer" checked="checked" />
        Microsoft SQL Server 2005 or above</label>
        <br />
        <span class="smText">This is the best performing version of Web Wiz Forums. With it's advanced paging it can easily cope with huge busy forums with over 100,000 members and millions of posts.</span></td>
    </tr>
    <tr>
      <td align="center"  class="tableRow"><img src="forum_images/MSSQL.gif" alt="Microsoft SQL Server 2000" width="64" height="64" /></td>
      <td colspan="2"  class="tableRow"><label>
        <input type="radio" name="DbType" id="DbType2" value="SQLServer2000" />
        Microsoft SQL Server 2000</label>
        <br />
        <span class="smText">The SQL Server 2000 version is without the advanced paging found in two  versions above. This version is able to cope very well with busy forums.</span></td>
    </tr>
    <tr>
      <td align="center"  class="tableRow"><img src="forum_images/MySQL.gif" alt="MySQL 4.1 or above" width="64" height="64" /></td>
      <td colspan="2"  class="tableRow"><label>
        <input type="radio" name="DbType" id="DbType3" value="mySQL" />
        MySQL 4.1 or above</label>
        <br />
        <span class="smText">mySQL is an excellent database for Web Wiz Forums. This version also includes the advanced paging found in the SQL Server 2005/2008 version  and is able to cope with huge busy forums with ease.</span></td>
    </tr>
    <tr>
      <td align="center"  class="tableRow"><img src="forum_images/MSAccess.gif" alt="Microsoft Access" width="64" height="64" /></td>
      <td colspan="2"  class="tableRow"><label>
        <input type="radio" name="DbType" id="DbType4" value="Access" />
        Microsoft Access</label>
        <br />
        <span class="smText">Although simple to setup, Access is only supported for test installations and not in the production environment. Access is a desktop database and is often unable to cope in a production environment leading to database corruption issues.</span> </td>
    </tr>
    <tr>
      <td colspan="3"  class="tableBottomRow">
        <table width="100%" border="0" cellspacing="2" cellpadding="0">
          <tr>
            <td width="52%"><input name="install" type="hidden" id="install" value="<% = Request.Form("install") %>" /></td>
            <td width="48%"><input type="submit" name="button" id="button" value="Next &gt;&gt;" /></td>
          </tr>
        </table></td>
    </tr>
  </table>
</form>
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
</body>
</html>
