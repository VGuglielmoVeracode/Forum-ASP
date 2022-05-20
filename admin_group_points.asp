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




      

'Read in the users details for the forum
intPointsTopic = IntC(Request.Form("TopicPoints"))
intPointsReply = IntC(Request.Form("ReplyPoints"))
intPointsAnswered = IntC(Request.Form("AnswerPoints"))
intPointsThanked = IntC(Request.Form("ThankedPoints"))
			

				
			



'If the user is changing tthe colours then update the database
If Request.Form("postBack") AND blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))	
	
	Call addConfigurationItem("Points_topic", intPointsTopic)
	Call addConfigurationItem("Points_reply", intPointsReply)
	Call addConfigurationItem("Points_answer", intPointsAnswered)
	Call addConfigurationItem("Points_thanked", intPointsThanked)
	
					
	'Update variables
	Application.Lock
	Application(strAppPrefix & "intPointsTopic") = Cint(intPointsTopic)
	Application(strAppPrefix & "intPointsReply") = Cint(intPointsReply)
	Application(strAppPrefix & "intPointsAnswered") = Cint(intPointsAnswered)
	Application(strAppPrefix & "intPointsThanked") = Cint(intPointsThanked)
	
	
	Application(strAppPrefix & "blnConfigurationSet") = false
	Application.UnLock
End If






'Initialise the SQL variable with an SQL statement to get the configuration details from the database
strSQL = "SELECT " & strDbTable & "SetupOptions.Option_Item, " & strDbTable & "SetupOptions.Option_Value " & _
"FROM " & strDbTable & "SetupOptions" &  strDBNoLock & " " & _
"ORDER BY " & strDbTable & "SetupOptions.Option_Item ASC;"

	
'Query the database
rsCommon.Open strSQL, adoCon

'Read in the forum from the database
If NOT rsCommon.EOF Then
	
	'Place into an array for performance
	saryConfiguration = rsCommon.GetRows()

	
	'Read in the colour info from the database
	intPointsTopic = CInt(getConfigurationItem("Points_topic", "numeric"))
	intPointsReply = CInt(getConfigurationItem("Points_reply", "numeric"))
	intPointsAnswered = CInt(getConfigurationItem("Points_answer", "numeric"))
	intPointsThanked = CInt(getConfigurationItem("Points_thanked", "numeric"))
	
	
End If


'Reset Server Objects
rsCommon.Close
Call closeDatabase()
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Group Points System</title>
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
  <h1>Group Point System</h1> 
  <br />
    <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
    <br />
    From here you can setup the points system. <br />
    <br />
    <table border="0" cellpadding="4" cellspacing="1" class="tableBorder">
     <tr>
      <td align="center" class="tableLedger">Group Points System</td>
     </tr>
     <tr>
      <td align="left" class="tableRow">The Point System allows members to build-up points within the forums depending on how many new topics they create, replies, etc.<br />
       <br />
The number of points a member has accumulated is also used within the Group Ladder System to allow members to move up through Ladder Groups depending on the points they have.</td>
     </tr>
    </table>
    <br />
</div>
<form action="admin_group_points.asp<% = strQsSID1 %>" method="post" name="frmConfiguration" id="frmConfiguration" onsubmit="return CheckForm();">
  <table border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
      <td colspan="2" class="tableLedger">Point  Settings</td>
    </tr>
    <tr>
     <td width="57%" class="tableRow">Create New Topic:<br />
      <span class="smText">This is number of points a member receives for creating a new Topic.</span></td>
     <td width="43%" valign="top" class="tableRow">
      <select name="TopicPoints" id="TopicPoints"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option<% If intPointsTopic = 1 Then Response.Write(" selected") %>>1</option>
       <option<% If intPointsTopic = 2 Then Response.Write(" selected") %>>2</option>
       <option<% If intPointsTopic = 3 Then Response.Write(" selected") %>>3</option>
       <option<% If intPointsTopic = 4 Then Response.Write(" selected") %>>4</option>
       <option<% If intPointsTopic = 5 Then Response.Write(" selected") %>>5</option>
       <option<% If intPointsTopic = 6 Then Response.Write(" selected") %>>6</option>
       <option<% If intPointsTopic = 7 Then Response.Write(" selected") %>>7</option>
       <option<% If intPointsTopic = 8 Then Response.Write(" selected") %>>8</option>
       <option<% If intPointsTopic = 9 Then Response.Write(" selected") %>>9</option>
       <option<% If intPointsTopic = 10 Then Response.Write(" selected") %>>10</option>
       <option<% If intPointsTopic = 15 Then Response.Write(" selected") %>>15</option>
       <option<% If intPointsTopic = 20 Then Response.Write(" selected") %>>20</option>
      </select> 
      points     
     </td>
    </tr>
    <tr>
     <td class="tableRow">Post Reply:<br />
      <span class="smText">This is number of points a member receives for posting a Reply.</span></td>
     <td valign="top" class="tableRow">
      <select name="ReplyPoints" id="ReplyPoints"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option<% If intPointsReply = 1 Then Response.Write(" selected") %>>1</option>
       <option<% If intPointsReply = 2 Then Response.Write(" selected") %>>2</option>
       <option<% If intPointsReply = 3 Then Response.Write(" selected") %>>3</option>
       <option<% If intPointsReply = 4 Then Response.Write(" selected") %>>4</option>
       <option<% If intPointsReply = 5 Then Response.Write(" selected") %>>5</option>
       <option<% If intPointsReply = 6 Then Response.Write(" selected") %>>6</option>
       <option<% If intPointsReply = 7 Then Response.Write(" selected") %>>7</option>
       <option<% If intPointsReply = 8 Then Response.Write(" selected") %>>8</option>
       <option<% If intPointsReply = 9 Then Response.Write(" selected") %>>9</option>
       <option<% If intPointsReply = 10 Then Response.Write(" selected") %>>10</option>
       <option<% If intPointsReply = 15 Then Response.Write(" selected") %>>15</option>
       <option<% If intPointsReply = 20 Then Response.Write(" selected") %>>20</option>
      </select> 
      points
     </td>
    </tr>
     <tr>
     <td class="tableRow">Post Answer/Resolution:<br />
      <span class="smText">This is number of points a member receives for posting a Reply that is set as an Answer or Resolution to the Topic.</span></td>
     <td valign="top" class="tableRow">
      <select name="AnswerPoints" id="AnswerPoints"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option<% If intPointsAnswered = 1 Then Response.Write(" selected") %>>1</option>
       <option<% If intPointsAnswered = 2 Then Response.Write(" selected") %>>2</option>
       <option<% If intPointsAnswered = 3 Then Response.Write(" selected") %>>3</option>
       <option<% If intPointsAnswered = 4 Then Response.Write(" selected") %>>4</option>
       <option<% If intPointsAnswered = 5 Then Response.Write(" selected") %>>5</option>
       <option<% If intPointsAnswered = 6 Then Response.Write(" selected") %>>6</option>
       <option<% If intPointsAnswered = 7 Then Response.Write(" selected") %>>7</option>
       <option<% If intPointsAnswered = 8 Then Response.Write(" selected") %>>8</option>
       <option<% If intPointsAnswered = 9 Then Response.Write(" selected") %>>9</option>
       <option<% If intPointsAnswered = 10 Then Response.Write(" selected") %>>10</option>
       <option<% If intPointsAnswered = 15 Then Response.Write(" selected") %>>15</option>
       <option<% If intPointsAnswered = 20 Then Response.Write(" selected") %>>20</option>
      </select> 
      points
     </td>
    </tr>
     <tr>
     <td class="tableRow">Thanked:<br />
      <span class="smText">This is number of points a member receives for being Thanked for a Post they have posted.</span></td>
     <td valign="top" class="tableRow">
      <select name="ThankedPoints" id="ThankedPoints"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option<% If intPointsThanked = 1 Then Response.Write(" selected") %>>1</option>
       <option<% If intPointsThanked = 2 Then Response.Write(" selected") %>>2</option>
       <option<% If intPointsThanked = 3 Then Response.Write(" selected") %>>3</option>
       <option<% If intPointsThanked = 4 Then Response.Write(" selected") %>>4</option>
       <option<% If intPointsThanked = 5 Then Response.Write(" selected") %>>5</option>
       <option<% If intPointsThanked = 6 Then Response.Write(" selected") %>>6</option>
       <option<% If intPointsThanked = 7 Then Response.Write(" selected") %>>7</option>
       <option<% If intPointsThanked = 8 Then Response.Write(" selected") %>>8</option>
       <option<% If intPointsThanked = 9 Then Response.Write(" selected") %>>9</option>
       <option<% If intPointsThanked = 10 Then Response.Write(" selected") %>>10</option>
       <option<% If intPointsThanked = 15 Then Response.Write(" selected") %>>15</option>
       <option<% If intPointsThanked = 20 Then Response.Write(" selected") %>>20</option>
      </select> 
      points
     </td>
    </tr>
    <tr align="center">
      <td colspan="2" class="tableBottomRow">
          <input type="hidden" name="postBack" value="true" />
          <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
          <input type="submit" name="Submit" value="Update Point Settings" />
          <input type="reset" name="Reset" value="Reset Form" />      </td>
    </tr>
  </table>
</form>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
