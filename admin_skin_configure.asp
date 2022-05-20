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


'Make sure this page is not cached
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"



Dim strForumName 		'Holds the forum name




'Read in the users details for the forum
If blnDemoMode = False Then
	strForumName = Request.Form("forumName")
	strTitleImage = Request.Form("titleImage")
	strCSSfile = Request.Form("CSS")
	strImagePath = Request.Form("imagePath")
	strNavSpacer = Request.Form("navSep")
	strHeader = Request.Form("Header")
	strFooter = Request.Form("Footer")
	blnShowHeaderFooter = BoolC(Request.Form("EnableHeaderFooter"))
	
End If



'If this is a pre-installed skin slection then setup the details
If Request.Form("Submit2") <> "" Then 
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
	
	'Read in the site name
	If blnDemoMode = False Then
		strForumName = Request.Form("forumName")
		strNavSpacer = Request.Form("navSep")
	Else
		strForumName = "Web Wiz Forums"
		strNavSpacer = " > "
	End If

	
	'Dark skin
	If Request.Form("skin") = "dark" Then
		If InStr(strTitleImage, "web_wiz_forums") Then strTitleImage = "forum_images/web_wiz_forums_black.png"
		strCSSfile = "css_styles/dark/"
		strImagePath = "forum_images/"
	
	'Vista Blue skin
	ElseIf Request.Form("skin") = "vistaBlue" Then
		If InStr(strTitleImage, "web_wiz_forums") Then strTitleImage = "forum_images/web_wiz_forums.png"
		strCSSfile = "css_styles/vista_blue/"
		strImagePath = "forum_images/"
		
	'Web Wiz skin
	ElseIf Request.Form("skin") = "webwiz" Then
		If InStr(strTitleImage, "web_wiz_forums") Then strTitleImage = "forum_images/web_wiz_forums.png"
		strCSSfile = "css_styles/web_wiz/"
		strImagePath = "forum_images/"
	
	'Classic skin	
	ElseIf Request.Form("skin") = "classic" Then
		If InStr(strTitleImage, "web_wiz_forums") Then strTitleImage = "forum_images/web_wiz_forums.png"
		strCSSfile = "css_styles/classic/"
		strImagePath = "forum_images/"	
		
	
	'Default skin	
	Else
		If InStr(strTitleImage, "web_wiz_forums") Then strTitleImage = "forum_images/web_wiz_forums.png"
		strCSSfile = "css_styles/default/"
		strImagePath = "forum_images/"
		
	End If
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
	
	Call addConfigurationItem("forum_name", Trim(Mid(strForumName, 1, 50)))
	Call addConfigurationItem("Title_image", Trim(Mid(strTitleImage, 1, 175)))
	Call addConfigurationItem("Skin_file", Trim(Mid(strCSSfile, 1, 100)))
	Call addConfigurationItem("Skin_image_path", Trim(Mid(strImagePath, 1, 100)))
	Call addConfigurationItem("Skin_nav_spacer", Mid(strNavSpacer, 1, 15))
	
	
	'Update variables
	Application.Lock
	Application(strAppPrefix & "strMainForumName") = strForumName
	Application(strAppPrefix & "strTitleImage") = strTitleImage
	Application(strAppPrefix & "strCSSfile") = strCSSfile
	Application(strAppPrefix & "strImagePath") = strImagePath
	Application(strAppPrefix & "strNavSpacer") = strNavSpacer
	
	
	
	'Empty the application level variable so that the changes made are seen in the main forum
	Application(strAppPrefix & "blnConfigurationSet") = false
	
	Application.UnLock

End If


'If the user is changing the details
If Request.Form("postBack") AND strForumName <> "" Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
	
	Call addConfigurationItem("forum_name", Trim(Mid(strForumName, 1, 50)))
	Call addConfigurationItem("Title_image", Trim(Mid(strTitleImage, 1, 175)))
	Call addConfigurationItem("Skin_file", Trim(Mid(strCSSfile, 1, 100)))
	Call addConfigurationItem("Skin_image_path", Trim(Mid(strImagePath, 1, 100)))
	Call addConfigurationItem("Skin_nav_spacer", Mid(strNavSpacer, 1, 15))
	Call addConfigurationItem("Header", strHeader)
	Call addConfigurationItem("Footer", strFooter)
	Call addConfigurationItem("Show_header_footer", blnShowHeaderFooter)
	
			
	
	'Update variables
	Application.Lock
	Application(strAppPrefix & "strMainForumName") = strForumName
	Application(strAppPrefix & "strTitleImage") = strTitleImage
	Application(strAppPrefix & "strCSSfile") = strCSSfile
	Application(strAppPrefix & "strImagePath") = strImagePath
	Application(strAppPrefix & "strNavSpacer") = strNavSpacer
	Application(strAppPrefix & "strHeader") = strHeader
	Application(strAppPrefix & "strFooter") = strFooter
	Application(strAppPrefix & "blnShowHeaderFooter") = CBool(blnShowHeaderFooter)
	
	
	
	'Empty the application level variable so that the changes made are seen in the main forum
	Application(strAppPrefix & "blnConfigurationSet") = false
	
	Application.UnLock
	
End If



'Initialise the SQL variable with an SQL statement to get the configuration details from the database
strSQL = "SELECT " & strDbTable & "SetupOptions.Option_Item, " & strDbTable & "SetupOptions.Option_Value " & _
"FROM " & strDbTable & "SetupOptions" &  strDBNoLock & " " & _
"ORDER BY " & strDbTable & "SetupOptions.Option_Item ASC;"

	
'Query the database
rsCommon.Open strSQL, adoCon


'Read in the forum colours from the database
If NOT rsCommon.EOF Then
	
	'Place into an array for performance
	saryConfiguration = rsCommon.GetRows()

	
	'Read in the colour info from the database
	strForumName = getConfigurationItem("forum_name", "string")
	strTitleImage = getConfigurationItem("Title_image", "string")
	strCSSfile = getConfigurationItem("Skin_file", "string")
	strImagePath = getConfigurationItem("Skin_image_path", "string")
	strNavSpacer = getConfigurationItem("Skin_nav_spacer", "string")
	strHeader = getConfigurationItem("Header", "string")
	strFooter = getConfigurationItem("Footer", "string")
	blnShowHeaderFooter = CBool(getConfigurationItem("Show_header_footer", "bool"))
	
End If

If strImagePath <> "" AND isNull(strImagePath) = False Then 
	If Mid(strImagePath, len(strImagePath), 1) <> "/" Then  strImagePath = strImagePath & "/"
End If
If strCSSfile <> "" AND isNull(strCSSfile) = False Then
	If Mid(strCSSfile, len(strCSSfile), 1) <> "/" Then  strCSSfile = strCSSfile & "/"
End If


'Reset Server Objects
rsCommon.Close
Call closeDatabase()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="includes/browser_page_encoding_inc.asp" -->
<title>Forum Skin Configuration</title>
<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write("<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Forums(TM) ver. " & strVersion & "" & _
vbCrLf & "Info: http://www.webwizforums.com" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>
<!-- Check the from is filled in correctly before submitting -->
<script  language="JavaScript" type="text/javascript">
<!-- Hide from older browsers...

//Function to check form is filled in correctly before submitting
function CheckForm () {

	if (document.frmSkinSetup.forumName.value==""){
		alert("Please enter the name of your forum");
		document.frmConfiguration.forumName.focus();
		return false;
	}
	
	//Check for a CSS file
	if (document.frmSkinSetup.CSS.value==""){
		alert("Please enter the a CSS file path");
		document.frmConfiguration.CSS.focus();
		return false;
	}
	
	//Check for a image path
	if (document.frmSkinSetup.imagePath.value==""){
		alert("Please enter the path to your forum images");
		document.frmConfiguration.imagePath.focus();
		return false;
	}
	return true;
}
// -->
</script>
<script language="JavaScript" src="includes/default_javascript_v9.js" type="text/javascript"></script>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
</head>
<body>
<!-- #include file="includes/admin_header_inc.asp" -->
   <h1>Forum  Skin Configuration</h1>
   <a href="admin_menu.asp">Control Panel Menu</a><br />
   <br />
   From here you can apply new skins to to change the look of Web Wiz Forums. <br />
   <br />
   <br />
   <table border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
     <td class="tableLedger">Copyright Notice</td>
    </tr>
    <tr>
     <td class="tableRow">The icons and images used within this software are copyright of '<a href="https://www.webwiz.net/" target="_blank">Web Wiz</a>', '<a href="http://www.awicons.com/stockicons/?discount=webwizguide" target="_blank">AW Icons</a>' and '<a href="http://www.glyphlab.com/" target="_blank">Glyph Lab</a>', if you wish to use any of the icons or images outside of the software package a license will need to be purchased from the copyright owner (This applies to ALL editions including the Free and Premium Editions).<br />
      <br />
     For licensing information please contact '<a href="https://www.webwiz.net/" target="_blank">Web Wiz</a>'. </td>
    </tr>
   </table>
   <br />
   <form action="admin_skin_configure.asp<% = strQsSID1 %>" method="post" name="frmSkinSetup" id="frmSkinSetup" onsubmit="return CheckForm();">
    <table border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
     <tr>
      <td colspan="2" class="tableLedger">Custom Skin Settings</td>
     </tr>
     <tr class="text">
      <td height="12" align="left" class="tableRow">Forum Name <br />
       <span class="smText">The name of your Forum pages eg. 'My Website Forum'.</span></td>
      <td height="12" valign="top" class="tableRow"><input name="forumName" type="text" id="forumName" value="<% = strForumName %>" size="30" maxlength="50"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />      </td>
     </tr>
     <tr class="text">
      <td height="2" align="left" class="tableRow">Forum Banner  Image <br />
       <span class="smText">This replaces the Web Wiz Forums banner logo that is shown on the top of each page. Your own web sites banner logo would be good to place here. </span><br /></td>
      <td height="2" valign="top" class="tableRow"><input type="text" name="titleImage" maxlength="175" value="<% = strTitleImage %>" size="30"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
     </tr>
     <tr>
      <td width="57%" class="tableRow">Skin Folder Name <br />
       <span class="smText">This is the folder name of where your CSS (Cascading Style Sheet) Files that are used to style Web Wiz Forums are stored </span></td>
      <td width="43%" valign="top" class="tableRow"><input name="CSS" type="text" id="CSS" value="<% = strCSSfile %>" size="30" maxlength="100"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
     </tr>
     <tr>
      <td class="tableRow">Image Folder Name <br />
       <span class="smText">This is the folder name of where your images for the skin you are using are stored. </span></td>
      <td valign="top" class="tableRow"><input type="text" name="imagePath" id="imagePath" maxlength="100" value="<% = strImagePath %>" size="30"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />      </td>
     </tr>
     <tr>
      <td class="tableRow">Bread Crumb Navigation Separator*<br />
        <span class="smText">This is the character that separates the separate links found in the bread crumb navigation found at the top of forum pages. </span></td>
      <td valign="top" class="tableRow"><input name="navSep" type="text" id="navSep" value="<% = strNavSpacer %>" size="7" maxlength="10"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
     </tr>
     
    
     
    
     
    <tr>
     <td colspan="2" class="tableLedger">Header</td>
    </tr>
    
    <tr>
     <td class="tableRow" width="50%">Enable Custom Header and Footer:<br />
      <span class="smText">This will enable the Custom Header and Footer set below.</span></td>
     <td valign="top" class="tableRow" width="50%">Yes
      <input type="radio" name="EnableHeaderFooter" value="True" <% If blnShowHeaderFooter = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="EnableHeaderFooter" value="False" <% If blnShowHeaderFooter = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    
    <tr>
     <td colspan="2" class="tableRow">
     If you have a custom header you wish to use, such as a website template enter the HTML below that you want displayed at the top of your forum*
      <br />
     <textarea name="Header" id="Header" rows="7" cols="100"><% = strHeader %></textarea>
     </td>
    </tr>
    
    <tr>
     <td colspan="2" class="tableLedger">Footer</td>
    </tr>
    <tr>
     <td colspan="2" class="tableRow">
     If you have a custom footer you wish to use, such as a website template enter the HTML below that you want displayed at the bottom of your forum*
     <br />
     <textarea name="Footer" id="Footer" rows="7" cols="100"><% = strFooter %></textarea>
     <br />
     <span class="smText">*Custom Headers or Footers may not be compatible and may require modification to correctly display or prevent conflicts with page elements and features of Web Wiz Forums pages.</span>
     </td>
    </tr>
     
     
     <tr align="center">
      <td colspan="2" class="tableBottomRow">
       <input type="hidden" name="postBack" value="true" />
       <input type="hidden" name="formID" id="formID1" value="<% = getSessionItem("KEY") %>" />
       <input name="Submit1" type="submit" id="Submit1" value="Update Skin" />
       <input type="reset" name="Reset" value="Reset Form" />      </td>
     </tr>
    </table>
   </form>
   <br />
   <form action="admin_skin_configure.asp<% = strQsSID1 %>" method="post" name="frmSkinSelect" id="frmSkinSelect">
    <table border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
     <tr>
      <td class="tableLedger">Pre-installed Skin Selection </td>
     </tr>
      <td class="tableRow"><span class="tableRow">
       <input name="skin" type="radio" value="default"<% If strCSSfile = "css_styles/default/" Then Response.Write(" checked=""checked""") %> />
       <strong>Default Skin</strong><br />
       <br />
       <img src="forum_images/default_skin_preview.jpg" alt="Default Skin Preview" width="400" height="234" border="0" class="imgBorder" /><br />
       <br />
       </span></td>
     </tr>
     <tr>
      <td class="tableRow"><input name="skin" type="radio" value="classic"<% If strCSSfile = "css_styles/classic/" Then Response.Write(" checked=""checked""") %> />
       <strong>Web Wiz Forums Classic Skin</strong><br />
       <br />
       <img src="forum_images/classic_skin_preview.jpg" alt="Web Wiz Forums Classic Skin" width="400" height="229" border="0" class="imgBorder" /><br />
       <br /></td>
     </tr>
     <tr>
      <td class="tableRow"><span class="tableRow">
       <input name="skin" type="radio" value="dark"<% If strCSSfile = "css_styles/dark/" Then Response.Write(" checked=""checked""") %> />
       <strong>Dark Skin</strong><br />
       <br />
       <img src="forum_images/dark_skin_preview.jpg" alt="Dark Skin Preview" width="400" height="233" border="0" class="imgBorder" /><br />
       <br />
       </span></td>
     </tr>
     <tr>
      <td class="tableRow"><span class="tableRow">
       <input name="skin" type="radio" value="webwiz"<% If strCSSfile = "css_styles/web_wiz/" Then Response.Write(" checked=""checked""") %> />
       <strong>Web Wiz  Skin</strong><br />
       <br />
       <img src="forum_images/web_wiz_skin_preview.jpg" alt="Web Wiz Skin Preview" width="400" height="229" border="0" class="imgBorder" /><br />
       <br />
       </span></td>
     </tr>
     <tr align="center">
      <td class="tableBottomRow">
       <input name="forumName" type="hidden" id="forumName2" value="<% = strForumName %>" />
       <input name="titleImage" type="hidden" id="titleImage2" value="<% = strTitleImage %>" />
       <input name="navSep" type="hidden" id="navSep2" value="<% = strNavSpacer %>" />
       <input type="hidden" name="formID" id="formID2" value="<% = getSessionItem("KEY") %>" />
       <input type="submit" name="Submit2" value="Select Pre-Installed Skin" />      </td>
     </tr>
    </table>
   </form>
   <br />
   <!-- #include file="includes/admin_footer_inc.asp" -->