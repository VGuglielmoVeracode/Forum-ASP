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
<title>Swear/Bad Word Filter</title>
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
<div align="center"><h1>Swear/Bad Word Filter</h1><br />
  <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
  <br />
  <span class="text">Use the form below to add new Swear Words, Bad Words, or Text that you do not want in your forum and their replacements. <br />
If someone then enters one of the 'Bad Words' into their Post, Private Message, Chat Room Message, etc. it will be replaced with the word in the 'Replace With' field.<br />
<br />
Regular Expressions are supported for </span>the Bad Word filed allowing you to use more powerful Word Searching to locate Bad Words.<br />
For example if you wanted to replace the whole word <strong>prat</strong> but not words that had <strong>prat</strong> in them like <strong>pratfull</strong> the you would use<strong> \bprat\b</strong> in the Bad Word field.<br />
Bad Word Searches are not case senstive.<br />
  <br />
  <form action="admin_bad_word_filter_add.asp<% = strQsSID1 %>" method="post" name="frmAddNewWords" id="frmAddNewWords">
   <table border="0" cellpadding="4" cellspacing="1" class="tableBorder" style="width: 450px;">
    <tr align="center">
     <td width="50%" align="left" class="tableLedger">Bad Word to Search For</td>
     <td width="50%" align="left" class="tableLedger">Replace With</td>
    </tr>
    <tr>
     <td width="50%" class="tableRow"><input name="badWord1" type="text" size="25" maxlength="49" />
     </td>
     <td width="50%" class="tableRow"><input name="replaceWord1" type="text" size="30" maxlength="49" />
     </td>
    </tr>
    <tr>
     <td width="50%" class="tableRow"><input name="badWord2" type="text" size="25" maxlength="49" />
     </td>
     <td width="50%" class="tableRow"><input name="replaceWord2" type="text" size="30" maxlength="49" />
     </td>
    </tr>
    <tr>
     <td class="tableRow"><input name="badWord3" type="text" size="25" maxlength="49" /></td>
     <td class="tableRow"><input name="replaceWord3" type="text" size="30" maxlength="49" /></td>
    </tr>
    <tr class="tableBottomRow">
     <td colspan="2" align="center"><input type="hidden" name="formID" id="formID1" value="<% = getSessionItem("KEY") %>" />
       <input name="Submit2" type="submit" value="Add New Bad Words To List" /></td>
    </tr>
   </table>
  </form>
  <br />
  <br />
  <span class="text"><span class="lgText">Remove Bad Words From List</span><b><br />
  </b><br />
  Place a tick in the checkbox for any bad words you wish to remove from the list <br />
  then click on the Delete Bad Words from List button.</span> <br />
  <br />
  <form action="admin_bad_word_filter_delete.asp<% = strQsSID1 %>" method="post" name="frmModerators" id="frmModerators">
    <table border="0" cellpadding="4" cellspacing="1" class="tableBorder" style="width: 450px;">
      <tr>
        <td width="18%" height="2" align="center" class="tableLedger">Delete</td>
        <td width="44%" height="2" align="center" class="tableLedger">Bad Word</td>
        <td width="38%" height="2" align="center" class="tableLedger">Replaced With</td>
      </tr><%
						
'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT " & strDbTable & "Smut.* FROM " & strDbTable & "Smut ORDER BY " & strDbTable & "Smut.Smut ASC;"
				
'Query the database
rsCommon.Open strSQL, adoCon

'Display the bad words       		
Do While NOT rsCommon.EOF		
			%>
      <tr>
        <td width="18%" height="24" align="center" class="tableRow"><input type="checkbox" name="chkWordID" value="<% = rsCommon("ID_no") %>" />        </td>
        <td width="44%" height="24" align="left" class="tableRow"><% = rsCommon("Smut") %>        </td>
        <td width="38%" height="24" align="left" class="tableRow"><% = rsCommon("Word_replace") %>        </td>
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
        <input name="Submit" type="submit" value="Delete Bad Words From List" /></td>
      </tr>
    </table>
  </form>
  <br />
  <br />
</div>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
