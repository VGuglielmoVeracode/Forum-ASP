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
blnGuestSessions = BoolC(Request.Form("GuestSID"))
strBoardMetaDescription = Trim(removeAllTags(Request.Form("metaDescription")))
strBoardMetaKeywords = Trim(removeAllTags(Request.Form("metaKeywords")))
blnDynamicMetaTags = BoolC(Request.Form("DynamicMetaTags"))
blnSearchEngineSessions = BoolC(Request.Form("SEsessions"))
blnNoFollowTagInLinks = BoolC(Request.Form("nofollow"))
blnSeoTitleQueryStrings = BoolC(Request.Form("SeoTitle"))
strStatsTrackingCode = Trim(Request.Form("statsTracking"))
blnUrlRewrite = BoolC(Request.Form("UrlRewrite"))

'If the user is changing tthe colours then update the database
If Request.Form("postBack") AND blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
	
	Call addConfigurationItem("Guest_SID", blnGuestSessions)
	Call addConfigurationItem("Meta_description", strBoardMetaDescription)
	Call addConfigurationItem("Meta_keywords", strBoardMetaKeywords)
	Call addConfigurationItem("Meta_tags_dynamic", blnDynamicMetaTags)
	Call addConfigurationItem("Search_eng_sessions", blnSearchEngineSessions)
	Call addConfigurationItem("Hyperlinks_nofollow", blnNoFollowTagInLinks)
	Call addConfigurationItem("SEO_title", blnSeoTitleQueryStrings)
	Call addConfigurationItem("Stats_tracking_code", strStatsTrackingCode)
	Call addConfigurationItem("URL_Rewriting", blnUrlRewrite)
	
					
	'Update variables
	Application.Lock
	Application(strAppPrefix & "blnGuestSessions") = CBool(blnGuestSessions)
	Application(strAppPrefix & "strBoardMetaDescription") = strBoardMetaDescription
	Application(strAppPrefix & "strBoardMetaKeywords") = strBoardMetaKeywords
	Application(strAppPrefix & "blnDynamicMetaTags") = CBool(blnDynamicMetaTags)
	Application(strAppPrefix & "blnSearchEngineSessions") = CBool(blnSearchEngineSessions)
	Application(strAppPrefix & "blnNoFollowTagInLinks") = CBool(blnNoFollowTagInLinks)
	Application(strAppPrefix & "blnSeoTitleQueryStrings") = CBool(blnSeoTitleQueryStrings)
	Application(strAppPrefix & "strStatsTrackingCode") = strStatsTrackingCode
	Application(strAppPrefix & "blnUrlRewrite") = CBool(blnUrlRewrite)
	
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
	blnGuestSessions = CBool(getConfigurationItem("Guest_SID", "bool"))
	strBoardMetaDescription = getConfigurationItem("Meta_description", "string")
	strBoardMetaKeywords = getConfigurationItem("Meta_keywords", "string")
	blnDynamicMetaTags = CBool(getConfigurationItem("Meta_tags_dynamic", "bool"))
	blnSearchEngineSessions = CBool(getConfigurationItem("Search_eng_sessions", "bool"))
	blnNoFollowTagInLinks = CBool(getConfigurationItem("Hyperlinks_nofollow", "bool"))
	blnSeoTitleQueryStrings = CBool(getConfigurationItem("SEO_title", "bool"))
	strStatsTrackingCode = getConfigurationItem("Stats_tracking_code", "string")
	blnUrlRewrite = CBool(getConfigurationItem("URL_Rewriting", "bool"))
End If


'Reset Server Objects
rsCommon.Close
Call closeDatabase()
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Search Engine Optimisation (SEO)</title>
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
  <h1> Search Engine Optimisation (SEO)</h1> 
  <br />
    <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
    <br />
    From here you can configure your forum for Search Engine Indexing.<br />
    <br />
</div>
<form action="admin_seo_settings.asp<% = strQsSID1 %>" method="post" name="frmConfiguration" id="frmConfiguration" onsubmit="return CheckForm();">
  <table border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
      <td colspan="2" class="tableLedger">Search Engine Optimisation (SEO) Settings</td>
    </tr>
    <tr>
      <td width="57%" class="tableRow">Forum Meta Description:<br />
      <span class="smText">This is the Meta Description of your forums. This is used by Search Engines when indexing your forums and is often displayed as a description of your forums in search results. <a href="https://www.webwiz.net/kb/seo/meta-tags-tutorial.htm" target="_blank" class="smLink">More Info &gt;&gt;</a></span></td>
      <td width="43%" valign="top" class="tableRow"><input name="metaDescription" type="text" id="metaDescription" value="<% = strBoardMetaDescription %>" size="50" maxlength="200"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
     <td class="tableRow">Meta Keywords:<br />
       <span class="smText">Keywords are used to improve Search Engine Ranking. Use up to 15 keywords, separated by commas, that are common to your forum. (eg. forum, company name, widgets) <a href="https://www.webwiz.net/kb/seo/meta_tags_tutorial.asp" target="_blank" class="smLink">More Info &gt;&gt;</a></span></td>
     <td valign="top" class="tableRow"><input type="text" name="metaKeywords" id="metaKeywords" maxlength="200" value="<% = strBoardMetaKeywords %>" size="50"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />     </td>
    </tr>
    <tr>
     <td class="tableRow">Dynamic Meta Tags:<br />
      <span class="smText">When enabled dynamic meta tags will be created for your Forums and Topics creating Meta Descriptions and Keywords on the fly which improve Search Engine Indexing and Ranking.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="DynamicMetaTags" value="True" <% If blnDynamicMetaTags = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="DynamicMetaTags" value="False" <% If blnDynamicMetaTags = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
     <td class="tableRow">Page Titles in URL:<br />
      <span class="smText">When used with the <a href="http://www.webwizforums.com/download/" target="_blank" class="smLink">URL Rewriting Add-On</a> this will allow page names to be rewritten as HTML pages that include the title or name of the page within the file name. When used without URL Rewriting it will add a 'title' Query String to the URL with the title or name of the page, this can help improve SEO.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="SeoTitle" value="True" <% If blnSeoTitleQueryStrings = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="SeoTitle" value="False" <% If blnSeoTitleQueryStrings = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
     <td class="tableRow">URL Rewriting:<br />
      <span class="smText">This is one of the best ways to improve SEO in your forum by turning dynamic page URL's in to Search Engine Friendly HTML URL's. <br /><strong>DO NOT ENABLE</strong> this option unless you have first installed the Web Wiz URL Rewriting Tool Add-on or you will get 404 Page Not Found Errors in your forum.  <a href="https://www.webwiz.net/web-wiz-forums/forum-downloads.htm" class="smLink" target="_blank">Download the Web Wiz Forums URL Rewrite Tool</a>.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="UrlRewrite" value="True" <% If blnUrlRewrite = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="UrlRewrite" value="False" <% If blnUrlRewrite = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
      <td class="tableRow">Create Sessions for Search Engines (NOT RECOMMENDED):<br />
      <span class="smText">It is recommended that you do <strong>NOT</strong> enable this feature. By keeping this feature disabled, known Search Engines will not have an SID number added to URL's, which improves Search Engine Indexing.</span></td>
      <td valign="top" class="tableRow">Yes
       <input type="radio" name="SEsessions" value="True" <% If blnSearchEngineSessions = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
        &nbsp;&nbsp;No
        <input type="radio" name="SEsessions" value="False" <% If blnSearchEngineSessions = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />      </td>
    </tr>
     <tr>
     <td class="tableRow">Guest Sessions:<br />
       <span class="smText">By disabling Guest Sessions it will improve SEO but will mean that those few Guest Users who do not have cookies enabled in their browser will have limited features available to them.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="GuestSID" value="True" <% If blnGuestSessions = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="GuestSID" value="False" <% If blnGuestSessions = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
     <td class="tableRow">Nofollow Hyperlinks:<br />
       <span class="smText">When enabled this will add 'nofollow' to hyperlinks posted by users. This tells search engines not to follow the link entered by the user. This is useful to put-off link spammers who spam forums with links to their sites to improve their own search engine ranking.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="nofollow" value="True" <% If blnNoFollowTagInLinks = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="nofollow" value="False" <% If blnNoFollowTagInLinks = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    
     <tr>
     <td valign="top" class="tableRow">Google Analytics or Stats Tracking Code:<br />
       <span class="smText">If you are using a Stats Services like <a href="http://www.google.com/analytics/" target="_blank" class="smText">Google Analytics</a> that requires that you add tracking code to your pages copy and paste the Tracking Code in to the textbox.</span></td>
     <td valign="top" class="tableRow">
      <textarea name="statsTracking" id="statsTracking" rows="6" cols="40"><% = strStatsTrackingCode %></textarea>
     </td>
    </tr>
    
    <tr align="center">
      <td colspan="2" class="tableBottomRow">
          <input type="hidden" name="postBack" value="true" />
          <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
          <input type="submit" name="Submit" value="Update SEO Settings" />
          <input type="reset" name="Reset" value="Reset Form" />      </td>
    </tr>
  </table>
</form>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
