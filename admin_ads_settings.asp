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
strForumHeaderAd = Trim(Request.Form("headerAdTxt"))
strForumPostAd = Trim(Request.Form("postAdTxt"))
strVigLinkKey = Trim(removeAllTags(Request.Form("VigLinkKey")))

'If VigLink is disabled
If Request.Form("VigLinkKeyDisable") = "Disabled" Then
	strVigLinkKey = "Disabled"
End If


'If the user is changing tthe colours then update the database
If Request.Form("postBack") AND blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
	
	Call addConfigurationItem("Forum_header_ad", strForumHeaderAd)
	Call addConfigurationItem("Forum_post_ad", strForumPostAd)
	Call addConfigurationItem("VigLink_key", strVigLinkKey)
	
					
	'Update variables
	Application.Lock
	Application(strAppPrefix & "strForumHeaderAd") = strForumHeaderAd
	Application(strAppPrefix & "strForumPostAd") = strForumPostAd
	Application(strAppPrefix & "strVigLinkKey") = strVigLinkKey
	
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
	strForumHeaderAd = getConfigurationItem("Forum_header_ad", "string")
	strForumPostAd = getConfigurationItem("Forum_post_ad", "string")
	strVigLinkKey = getConfigurationItem("VigLink_key", "string")
End If


'Reset Server Objects
rsCommon.Close
Call closeDatabase()
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Affiliate and Text/Banner Ads</title>
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
  <h1>Affiliate and Text/Banner Ads</h1> 
  <br />
    <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
    <br />
    From here you can Monetize your forum content by affiliating links using VigLink and/or add Banner and Text Ads to your Forum.<br />
    <br />
</div>
<form action="admin_ads_settings.asp<% = strQsSID1 %>" method="post" name="frmConfiguration" id="frmConfiguration" onsubmit="return CheckForm();">
  <table border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
      <td colspan="2" class="tableLedger">Affiliate and Text/Banner Ads</td>
    </tr>
    <tr>
      <td width="57%" class="tableRow"><a href="http://www.viglink.com/?vgref=12412" target="_blank">VigLink Key</a>:<br />
      <span class="smText"><a href="http://www.viglink.com/?vgref=12412" target="_blank" class="smText">VigLink</a> is one of the simplest ways to monetize your forums content.
      <br />
      
      
      
      
      <br />
      To activate VigLink in your forum <a href="http://www.viglink.com/?vgref=12412" target="_blank" class="smText">sign to VigLink</a> and then copy and paste the API Key found on your VigLink Account page (eg. a39e8374ae6f....7fbe31) in to the text area on the right.
      <br /><br />
      VigLink works by invisibly monitoring links within the content of your forum and when a link is clicked it silently affiliates the link. For example, a member posts about an item on eBay, later someone clicks that links and buys a product from eBay and VigLink pays you the commission. </span>
      <br /><br />
      </td>
      <td width="43%" class="tableRow"><input name="VigLinkKey" type="text" id="VigLinkKey" value="<% If NOT strVigLinkKey = "Disabled" Then Response.Write(strVigLinkKey) %>" size="40" maxlength="50"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
       <br />
       
        <label for="SfsTerms"><input type="checkbox" name="VigLinkKeyDisable" id="VigLinkKeyDisable" value="Disabled" <% If strVigLinkKey = "Disabled" Then Response.Write(" checked=""checked""") %> />
       
       Turn Off VigLink Ads</label>
       
       <br /> <br />
       <a href="http://www.viglink.com/?vgref=12412" target="_blank"><img src="<% = strImagePath %>VigLink.png" alt="VigLink" border="0" /></a>
       <br /> <br />
      </td>
    </tr>
    
     <tr>
     <td valign="top" class="tableRow">Header Ad Code:<br />
       <span class="smText">Copy and paste your Google Adsense Code or other Ad Code in to the text area on the right to display Ads at the top of your Forum.</span></td>
     <td valign="top" class="tableRow">
      <textarea name="headerAdTxt" id="headerAdTxt" rows="6" cols="40"><% = strForumHeaderAd %></textarea>
     </td>
    </tr>
    
    <tr>
     <td valign="top" class="tableRow">Post Ad Code:<br />
       <span class="smText">Copy and paste your Google Adsense Code or other Ad Code in to the text area on the right to display Ads between the first and secound posts within Forum Topics.</span></td>
     <td valign="top" class="tableRow">
      <textarea name="postAdTxt" id="postAdTxt" rows="6" cols="40"><% = strForumPostAd %></textarea>
     </td>
    </tr>
    
    <tr align="center">
      <td colspan="2" class="tableBottomRow">
          <input type="hidden" name="postBack" value="true" />
          <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
          <input type="submit" name="Submit" value="Update Ad Settings" />
          <input type="reset" name="Reset" value="Reset Form" />      </td>
    </tr>
  </table>
</form>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
