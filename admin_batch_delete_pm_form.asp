<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="admin_common.asp" -->
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




'Set the response buffer to true
Response.Buffer = True 


'Reset Server Objects
Call closeDatabase()
%>  
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title>Batch Delete Private Messages</title>

<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write("<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Forums(TM) ver. " & strVersion & "" & _
vbCrLf & "Info: http://www.webwizforums.com" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->" & vbCrLf & vbCrLf)
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>
  
<!-- #include file="includes/admin_header_inc.asp" -->
<div align="center">
 <h1>Batch Delete Private Messages</h1><br />
  <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
  <br />
  If you find the forum starts running a bit slow it maybe worth cleaning the database out by deleting Private Messages.<br />
  <br />
  Select the Private Messages you want deleted by the date the Private Message was sent.<br />
  <br />
  </p>
</div>
<form method="post" name="frmDeletePM" action="admin_batch_delete_pm.asp<% = strQsSID1 %>" onSubmit="return confirm('Are you sure you want to delete these Private Messages?\n\nOnce the Private Messages are deleted they will be lost forever.')">
  <table align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
      <td class="tableLedger">Delete Private Messages older than</td>
    </tr>
    <tr>
      <td align="left" class="tableRow">
      	 <select name="days" id="days">
          <option value="0">Now</option>
          <option value="7">1 Week</option>
          <option value="14">2 Weeks</option>
          <option value="31">1 Month</option>
          <option value="62">2 Months</option>
          <option value="124">4 Months</option>
          <option value="182" selected>6 Months</option>
          <option value="279">9 Months</option>
          <option value="365">1 Year</option>
          <option value="730">2 Years</option>
      </select></td>
    </tr>
    <tr class="tableBottomRow">
      <td align="center">
      	<input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
      	<input name="Submit" type="submit" value="Delete Private Messages" /></td>
    </tr>
  </table>
 </form>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->