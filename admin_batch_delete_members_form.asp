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
<title>Batch Delete Members</title>
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
 <h1>Batch Delete Members</h1><br />
    <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
    <br />
    If you find the forum starts running a bit slow it maybe worth cleaning the database out by deleting Members who have never posted.<br />
    <br />
    Select the length of time it has been since the <strong>non-posting</strong> members signed up on.<br />
    <br />
    </p>
</div>
<form action="admin_batch_delete_members.asp<% = strQsSID1 %>" method="post" name="frmDelete" id="frmDelete" onsubmit="return confirm('Are you sure you want to delete these Members?\n\nOnce the Members are deleted they will be lost forever.')">
  <table border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
      <td height="2" align="left" class="tableLedger">Delete Members that have never posted that registered over</td>
    </tr>
    <tr>
      <td class="tableRow"><select name="days">
          <option value="0">Now</option>
          <option value="7">1 Week</option>
          <option value="14">2 Weeks</option>
          <option value="31" selected>1 Month</option>
          <option value="62">2 Months</option>
          <option value="124">4 Months</option>
          <option value="182">6 Months</option>
          <option value="279">9 Months</option>
          <option value="365">1 Year</option>
          <option value="730">2 Years</option>
        </select>
      </td>
    </tr>
    <tr>
      <td  height="12" align="left" class="tableLedger">Select the type of members to delete</td>
    </tr>
    <tr>
      <td  height="12" align="left" class="tableRow"><input name="unactive" type="radio" value="false" checked="checked" />
        All member accounts<br />
        <input type="radio" name="unactive" value="true" />
        Un-activated member accounts</td>
    </tr>
    <tr class="tableBottomRow">
      <td align="center">
      	<input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
      	<input name="Submit" type="submit" value="Delete Members" /></td>
    </tr>
  </table>
</form>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
