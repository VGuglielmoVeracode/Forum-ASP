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




'Set the response buffer to true as we maybe redirecting
Response.Buffer = True

'If in demo mode redirect
If blnDemoMode Then
	Call closeDatabase()
	Response.Redirect("admin_web_wiz_forums_premium.asp" & strQsSID1)
End If
 

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title>SPAM Filters</title>
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
<div align="center"><h1>SPAM Filters</h1><br />
  <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
  <br />
  <span class="text">Use the form below to add new SPAM that you wish to filter. <br />
If the SPAM is detected in a Forum Post being submitted then the selected action will be taken. If SPAM is detected in a Private Message then it will be rejected.<br />
<br />
Regular Expressions are supported for the SPAM Filtering allowing you to use more powerful Word Searching to locate SPAM.<br />
For example if you wanted to search for the whole word <strong>spam</strong> but not words that had <strong>spam</strong> in them like <strong>spammer</strong> the you would use<strong> \bspam\b</strong> in the SPAM field.<br />
The SPAM filter is not case senstive.<br /></span>
  <br />
  <form action="admin_spam_filter_add.asp<% = strQsSID1 %>" method="post" name="frmSpamNewWords" id="frmSpamNewWords">
   <table border="0" cellpadding="4" cellspacing="1" class="tableBorder" style="width: 450px;">
    <tr align="center">
     <td width="50%" align="left" class="tableLedger">SPAM to Search For</td>
     <td width="50%" align="left" class="tableLedger">Forum Post Action</td>
    </tr>
    <tr>
     <td width="50%" class="tableRow"><input name="spamWord1" type="text" size="40" maxlength="200" />
     </td>
     <td width="50%" class="tableRow">
     	<select name="spamAction1">
     	 <option>Reject</option>
     	 <option>Require Approval</option>
        </select>
     </td>
    </tr>
    <tr>
     <td width="50%" class="tableRow"><input name="spamWord2" type="text" size="40" maxlength="200" />
     </td>
     <td width="50%" class="tableRow">
     	<select name="spamAction2">
     	 <option>Reject</option>
     	 <option>Require Approval</option>
        </select>
     </td>
    </tr>
    <tr>
     <td class="tableRow"><input name="spamWord3" type="text" size="40" maxlength="200" /></td>
     <td class="tableRow">
        <select name="spamAction3">
     	 <option>Reject</option>
     	 <option>Require Approval</option>
        </select>
     </td>
    </tr>
    <tr class="tableBottomRow">
     <td colspan="2" align="center"><input type="hidden" name="formID" id="formID1" value="<% = getSessionItem("KEY") %>" />
       <input name="Submit2" type="submit" value="Add New SPAM To List" /></td>
    </tr>
   </table>
  </form>
  <br />
  <br />
  <span class="text"><span class="lgText">Remove SPAM Filters From List</span><b><br />
  </b><br />
  Place a tick in the checkbox for any SPAM Filters you wish to remove from the list <br />
  then click on the Delete SPAM Filters from List button.</span> <br />
  <br />
  <form action="admin_spam_filter_delete.asp<% = strQsSID1 %>" method="post" name="frmSPAM" id="frmSPAM">
    <table border="0" cellpadding="4" cellspacing="1" class="tableBorder" style="width: 550px;">
      <tr>
        <td width="10%" height="2" align="center" class="tableLedger">Delete</td>
        <td width="60%" height="2" align="center" class="tableLedger">SPAM</td>
        <td width="20%" height="2" align="center" class="tableLedger">Forum Post Action</td>
      </tr><%
						
'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT " & strDbTable & "Spam.* FROM " & strDbTable & "Spam ORDER BY " & strDbTable & "Spam.Spam ASC;"
				
'Query the database
rsCommon.Open strSQL, adoCon

'Display the spam filters 		
Do While NOT rsCommon.EOF		
			%>
      <tr>
        <td width="18%" height="24" align="center" class="tableRow"><input type="checkbox" name="chkSpamdID" value="<% = rsCommon("Spam_ID") %>" /></td>
        <td width="44%" height="24" align="left" class="tableRow"><% = rsCommon("Spam") %></td>
        <td width="38%" height="24" align="left" class="tableRow"><% = rsCommon("Spam_Action") %></td>
      </tr><%
		
	'Move to the next record in the database
	rsCommon.MoveNext
	
'Loop back round   	
Loop
	
'Reset server variable
rsCommon.Close
Call closeDatabase()
%>
      <tr class="tableBottomRow">
        <td colspan="3" align="center">
        <input type="hidden" name="formID" id="formID2" value="<% = getSessionItem("KEY") %>" />
        <input name="Submit" type="submit" value="Delete SPAM Filters From List" /></td>
      </tr>
    </table>
  </form>
  <br />
  <br />
</div>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
