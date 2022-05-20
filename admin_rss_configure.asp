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
blnRSS = BoolC(Request.Form("RSS"))
intRssTimeToLive = IntC(Request.Form("TTL"))
intRSSmaxResults = IntC(Request.Form("MaxResults"))





'If the user is changing tthe colours then update the database
If Request.Form("postBack") AND blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
	
	Call addConfigurationItem("RSS", blnRSS)
	Call addConfigurationItem("RSS_TTL", intRssTimeToLive)
	Call addConfigurationItem("RSS_max_results", intRSSmaxResults)

					
	'Update variables
	Application.Lock
	Application(strAppPrefix & "blnRSS") = CBool(blnRSS)
	Application(strAppPrefix & "intRssTimeToLive") = CInt(intRssTimeToLive)
	Application(strAppPrefix & "intRSSmaxResults") = CInt(intRSSmaxResults)
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
	blnRSS = CBool(getConfigurationItem("RSS", "bool"))
	intRssTimeToLive = CInt(getConfigurationItem("RSS_TTL", "numeric"))
	intRSSmaxResults = CInt(getConfigurationItem("RSS_max_results", "numeric"))
End If


'Reset Server Objects
rsCommon.Close
Call closeDatabase()
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>RSS Feed Settings</title>
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
  <h1> RSS Feed Settings</h1> 
  <br />
    <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
    <br />
    From here you can configure RSS for your forum.
    </p>
    <br />
    <br />
   An RSS (Really Simple Syndication)  Feed will allow users to subscribe to the RSS Feed to be notified of new posts and calendar events.<br />
   Support for RSS Feeds is available in Windows Vista SideBar, Internet Explorer 7, Firefox, Thunderbird, Outlook, and many other programs.<br />
   RSS Feeds are public feeds without any type of authentication mechanism, for this reason only public Topics, Posts, Calendar Events will be syndicated.<br />
    <br />
</div>
<form action="admin_rss_configure.asp<% = strQsSID1 %>" method="post" name="frmConfiguration" id="frmConfiguration" onsubmit="return CheckForm();">
  <table border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
      <td colspan="2" class="tableLedger">RSS Feed Settings</td>
    </tr>
    <tr>
     <td width="57%" class="tableRow">RSS Feeds:<br />
       <span class="smText">This enables RSS Feeds for Forums, Topics, and Calendar Events.<br />
       </span></td>
     <td width="43%" valign="top" class="tableRow">Yes
      <input type="radio" name="RSS" value="True" <% If blnRSS = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
&nbsp;&nbsp;No
<input type="radio" name="RSS" value="False" <% If blnRSS = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
     <tr>
     <td class="tableRow">Time To Live:<br />
       <span class="smText">This is amount of time for the RSS Feed to live, before the Clients RSS Reader requests an update. If this is set to low you may consume to much bandwidth, to high and the RSS News Reader may not be updated fast enough.</span></td>
     <td valign="top" class="tableRow"><select name="TTL"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option<% If intRssTimeToLive = 10 Then Response.Write(" selected") %> value="10">10 minutes</option>
       <option<% If intRssTimeToLive = 30 Then Response.Write(" selected") %> value="30">30 minutes</option>
       <option<% If intRssTimeToLive = 60 Then Response.Write(" selected") %> value="60">1 hour</option>
       <option<% If intRssTimeToLive = 120 Then Response.Write(" selected") %> value="120">2 hours</option>
       <option<% If intRssTimeToLive = 180 Then Response.Write(" selected") %> value="180">3 hours</option>
       <option<% If intRssTimeToLive = 240 Then Response.Write(" selected") %> value="240">4  hours</option>
       <option<% If intRssTimeToLive = 300 Then Response.Write(" selected") %> value="300">5 hours</option>
       <option<% If intRssTimeToLive = 360 Then Response.Write(" selected") %> value="360">6 hours</option>
       <option<% If intRssTimeToLive = 720 Then Response.Write(" selected") %> value="720">12 hours</option>
     </select></td>
    </tr>
     <tr>
     <td class="tableRow">Maximum Number of Entries:<br />
       <span class="smText">This is maximum number of entries to display in RSS Feeds at one time. For example if you select 10 it will display the last 10 Topics, Posts, or Calender Events, depending on which RSS Feed the user is subscribed to.</span></td>
     <td valign="top" class="tableRow"><select name="MaxResults"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option<% If intRSSmaxResults = 2 Then Response.Write(" selected") %>>2</option>
       <option<% If intRSSmaxResults = 3 Then Response.Write(" selected") %>>3</option>
       <option<% If intRSSmaxResults = 4 Then Response.Write(" selected") %>>4</option>
       <option<% If intRSSmaxResults = 5 Then Response.Write(" selected") %>>5</option>
       <option<% If intRSSmaxResults = 6 Then Response.Write(" selected") %>>6</option>
       <option<% If intRSSmaxResults = 7 Then Response.Write(" selected") %>>7</option>
       <option<% If intRSSmaxResults = 8 Then Response.Write(" selected") %>>8</option>
       <option<% If intRSSmaxResults = 9 Then Response.Write(" selected") %>>9</option>
       <option<% If intRSSmaxResults = 10 Then Response.Write(" selected") %>>10</option>
       <option<% If intRSSmaxResults = 15 Then Response.Write(" selected") %>>15</option>
       <option<% If intRSSmaxResults = 20 Then Response.Write(" selected") %>>20</option>
       <option<% If intRSSmaxResults = 25 Then Response.Write(" selected") %>>25</option>
       <option<% If intRSSmaxResults = 30 Then Response.Write(" selected") %>>30</option>
       <option<% If intRSSmaxResults = 40 Then Response.Write(" selected") %>>40</option>
       <option<% If intRSSmaxResults = 50 Then Response.Write(" selected") %>>50</option>
       <option<% If intRSSmaxResults = 75 Then Response.Write(" selected") %>>75</option>
       <option<% If intRSSmaxResults = 100 Then Response.Write(" selected") %>>100</option>
     </select></td>
    </tr>
    <tr align="center">
      <td colspan="2" class="tableBottomRow">
          <input type="hidden" name="postBack" value="true" />
          <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
          <input type="submit" name="Submit" value="Update RSS Settings" />
          <input type="reset" name="Reset" value="Reset Form" />      </td>
    </tr>
  </table>
</form>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
