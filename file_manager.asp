<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="functions/functions_upload.asp" -->
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


Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"



'Dimension veriables
Dim strMode 			'Holds the page mode (eg admin)
Dim lngUserProfileID		'Holds the profile ID of the user
Dim dblFileSpaceUsed		'Holds the amount of file space used
Dim blnUploadFolderExsist	'Set to true if the user has an upload folder
Dim strFileName
Dim objFSO
Dim objFSOfile
Dim strFileType
Dim intFileSize
Dim strFileExtension
Dim blnFileUploadDetails
Dim strXID



blnUploadFolderExsist = False
blnFileUploadDetails = False



'If the user is not allowed kick 'em
If bannedIP() OR  blnActiveMember = False OR blnBanned OR intGroupID = 2 OR (blnAttachments = false AND blnImageUpload = false) Then
	
	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If




'Read in the mode of the page
strMode = Trim(Mid(Request.QueryString("M"), 1, 1))

'If this is not an admin but in admin mode then see if the user is a moderator
If blnAdmin = False AND strMode = "A" AND blnModeratorProfileEdit Then
	
	'Initalise the strSQL variable with an SQL statement to query the database
        strSQL = "SELECT " & strDbTable & "Permissions.Moderate " & _
        "FROM " & strDbTable & "Permissions " & _
        "WHERE (" & strDbTable & "Permissions.Group_ID = " & intGroupID & " OR " & strDbTable & "Permissions.Author_ID = " & lngLoggedInUserID & ") AND  " & strDbTable & "Permissions.Moderate=" & strDBTrue & ";"
               

        'Query the database
         rsCommon.Open strSQL, adoCon

        'If a record is returned then the user is a moderator in one of the forums
        If NOT rsCommon.EOF Then blnModerator = True

        'Clean up
        rsCommon.Close
End If





'Get the user ID of the memebr being edited by the admin and set to the users upload directory
If (blnAdmin OR (blnModerator AND LngC(Request.QueryString("PF")) > 2)) AND strMode = "A" Then
	
	lngUserProfileID = LngC(Request.QueryString("PF"))
	strUploadFilePath = strUploadOriginalFilePath & "/" & lngUserProfileID

'Get the logged in ID number
Else
	lngUserProfileID = lngLoggedInUserID
End If



'Check the user has an upload folder, if they have uploaded annything
blnUploadFolderExsist = userUploadFolder(strUploadFilePath)

'If the uopload folder exsists get the size
If blnUploadFolderExsist Then

	'Get the amount if file space used by this person
	dblFileSpaceUsed = folderSize(strUploadFilePath)
Else
	dblFileSpaceUsed = 0

End If






'If an upload has occured display trhe file details
If Request.QueryString("UL") = "true" AND Request.QueryString("FN") <> "" Then

	'Read in the file name
	strFileName = Trim(Request.QueryString("FN"))
	
	'Filer for malicous code
	strFileName = removeAllTags(strFileName)

	'If the user has an upload folder display the contents
	If blnUploadFolderExsist Then
		
		'Create the file system object
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		
		'Check to make sure file exsists
		If objFSO.FileExists(Server.MapPath(strUploadFilePath) & "\" & strFileName) Then
			
			'Create a file object with the file details
			Set objFSOfile = objFSO.GetFile(Server.MapPath(strUploadFilePath) & "\" & strFileName)
			
			'Read in the file details
			strFileName = objFSOfile.Name
			strFileType = objFSOfile.Type
			intFileSize = CInt(objFSOfile.Size / 1024)
			strFileExtension = LCase(objFSO.GetExtensionName(objFSOfile.Name))
			
			'Set to true
			blnFileUploadDetails = True
			
		End If
	End If
End If



'Get session key
strXID = getSessionItem("KEY")


'If active users is enabled update the active users application array
If blnActiveUsers Then
	'Call active users function
	saryActiveUsers = activeUsers("", strTxtFileManager, "file_manager.asp", 0)
End If

'Clean up
Call closeDatabase()


'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtFileManager


%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title><% = strTxtFileManager %></title>

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

<script language="JavaScript">

//delete file
function deleteFile(deleteFileName) {
	//check for file
	if (deleteFileName == ''){
		alert('<% = strTxtNoFileSelected %>');
		return;
	}
	//confirm file deletion
	if (confirm('<% = strTxtAreYouSureDeleteFile %> \'' + deleteFileName + '\'?')){
		self.location.href = 'file_delete.asp?fileName=' + escape(deleteFileName) + '<% If strMode = "A" Then Response.Write("&PF=" & lngUserProfileID & "&M=A") %>&XID=<% = strXID & strQsSID2 %>';
		return true;
	}
}
<%

'If a file has been uploaded preview it
If blnFileUploadDetails Then
	
%>

//Function to preview image
function uploadPreview(){
	//Get the file extension to check
	var extension = '<% = strFileName %>';
	extension = extension.slice(extension.lastIndexOf('.'), extension.length).toLowerCase();

	//Display file if of the correct type
	if (extension==".txt" 
	 || extension==".text" 
	 || extension==".htm" 
	 || extension==".html" 
	 || extension==".gif" 
	 || extension==".jpg" 
	 || extension==".jpeg" 
	 || extension==".png" 
	 || extension==".bmp"){
		try {
			document.getElementById('prevWindow').contentWindow.location.href = '<% = Replace(strUploadFilePath, "\", "/", 1, -1, 1) %>/<% = strFileName %>';	
		}catch(exception){
		}
	
	}else{
		document.getElementById("prevWindow").contentWindow.location.href="RTE_popup_link_preview.asp?b=0<% = strQsSID2 %>";
	
	}
	
	self.parent.document.getElementById('fileNameDisplay').innerHTML = '<% = strFileName %>';
	self.parent.document.getElementById('fileName').innerHTML = '<% = strFileName %>';
	self.parent.document.getElementById('fileType').innerHTML = '<% = strFileType %>';
	self.parent.document.getElementById('fileSize').innerHTML = '<% = intFileSize %>' + 'KB';
	self.parent.document.getElementById('fileDownload').innerHTML = '<a href="<% = Replace(strUploadFilePath, "\", "/", 1, -1, 1) %>/<% = strFileName %>" target="_blank"><% = strFileName %></a>';
}
<%

End If
%>
</script>

<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />  	
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% = strTxtFileManager %></h1></td>
</tr>
</table>
<br />
<table class="basicTable" cellspacing="0" cellpadding="0" align="center"> 
 <tr> 
  <td class="tabTable">
   <a href="member_control_panel.asp<% If strMode = "A" Then Response.Write("?PF=" & lngUserProfileID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>" title="<% = strTxtControlPanel %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>member_control_panel.<% = strForumImageType %>" border="0" alt="<% = strTxtControlPanel %>" /> <% = strTxtControlPanel %></a>
   <a href="register.asp<% If strMode = "A" Then Response.Write("?PF=" & lngUserProfileID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>" title="<% = strTxtProfile2 %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>profile.<% = strForumImageType %>" border="0" alt="<% = strTxtProfile2 %>" /> <% = strTxtProfile2 %></a><%
 
If blnEmail Then

%>
   <a href="email_notify_subscriptions.asp<% If strMode = "A" Then Response.Write("?PF=" & lngUserProfileID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>" title="<% = strTxtSubscriptions %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>subscriptions.<% = strForumImageType %>" border="0" alt="<% = strTxtSubscriptions %>" /> <% = strTxtSubscriptions %></a><%
End If


'Only disply other links if not in admin mode
If strMode <> "A" AND blnActiveMember AND blnPrivateMessages Then

%>
   <a href="pm_buddy_list.asp<% = strQsSID1 %>" title="<% = strTxtBuddyList %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>buddy_list.<% = strForumImageType %>" border="0" alt="<% = strTxtBuddyList %>" /> <% = strTxtBuddyList %></a><%

End If


'If the user is user is using a banned IP redirect to an error page
If blnAttachments OR blnImageUpload Then

%>
   <a href="file_manager.asp<% If strMode = "A" Then Response.Write("?PF=" & lngUserProfileID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>" title="<% = strTxtFileManager %>" class="tabButtonActive">&nbsp;<img src="<% = strImagePath %>file_manager.<% = strForumImageType %>" border="0" alt="<% = strTxtFileManager %>" /> <% = strTxtFileManager %></a><%

End If



'If in demo mode or alloted space is 0 set the file upload to full (prevents divsion errors later)
If blnDemoMode OR intUploadAllocatedSpace = 0 Then
	dblFileSpaceUsed = 0.001
	intUploadAllocatedSpace = 0.001
	
End If
%>
  </td>
 </tr>
</table>
<br />
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
 <tr class="tableLedger">
  <td colspan="2" align="left"><% = strTxtFileManager %></td>
 </tr>
 <tr class="tableRow">
  <td width="10%" valign="top">
   <div style="float:left;"><strong><% = strTxtFileExplorer %></strong></div>
   <div style="float:right;"><img src="<% = strImagePath %>file_upload.<% = strForumImageType %>" width="16" height="16" alt="<% = strTxtNewUpload %>" title="<% = strTxtNewUpload %>" /> <a href="file_upload.asp<% If strMode = "A" Then Response.Write("?PF=" & lngUserProfileID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>" title="<% = strTxtNewUpload %>" class="fileManLink"><% = strTxtNewUpload %></a>&nbsp;&nbsp;&nbsp;</div>
   <br />
   <div style="padding:8px;">
    <iframe src="file_browser.asp<% If strMode = "A" Then Response.Write("?PF=" & lngUserProfileID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>" id="fileWindow" width="333px" height="500px" style="border: #A5ACB2 1px solid;background-color: #FFFFFF;"></iframe>
   </div>
  </td>
  <td width="90%" valign="top">
   <strong><% = strTxtAllocatedFileSpace %></strong>
   <br />
    <table cellspacing="0" cellpadding="3" width="400">
     <tr class="tableRow"><td colspan="3"><% = strTxtYouHaveUsed & " " & dblFileSpaceUsed & "MB " & strTxtFromYour & " " & intUploadAllocatedSpace & "MB " & strTxtOfAllocatedFileSpace %></td></tr>
     <tr class="tableRow"><td colspan="3"><img src="<% = strImagePath %>bar_graph_image.gif" width="<% = FormatPercent((dblFileSpaceUsed / intUploadAllocatedSpace), 0) %>" height="13"  title="<% = strTxtYourFileSpaceIs & " " & FormatPercent((dblFileSpaceUsed / intUploadAllocatedSpace), 0) & " " & strTxtFull %>" /></td></tr>
     <tr class="tableRow">
      <td width="30%" class="smText">0MB</td>
      <td width="41%" align="center" class="smText"><% = FormatNumber(DblC(intUploadAllocatedSpace / 2), 1) %>MB</td>
      <td width="29%" align="right" class="smText"><% = intUploadAllocatedSpace %>MB</td>
     </tr>
    </table>
   <br />
   <strong><% = strTxtFileProperties %></strong>
   <table width="100%" border="0" cellspacing="4" cellpadding="0">
    <tr>
     <td align="right" width="15%"><% = strTxtFileName %>: </td>
     <td width="85%"><span id="fileNameDisplay"></span><input type="hidden" name="fileName" id="fileName" /></td>
    </tr>
    <tr>
     <td align="right"><% = strTxtFileSize %>: </td>
     <td><span id="fileSize"></span></td>
    </tr>
    <tr>
     <td align="right"><% = strTxtFileType %>: </td>
     <td><span id="fileType"></span></td>
    </tr>
    <tr>
     <td align="right" nowrap><% = strTxtDownloadFile %>: </td>
     <td><span id="fileDownload"></span></td>
    </tr>
    <tr>
     <td colspan="2"><br />
     <img src="<% = strImagePath %>file_upload.<% = strForumImageType %>" width="16" height="16" alt="<% = strTxtNewUpload %>" title="<% = strTxtNewUpload %>" /> <a href="file_upload.asp<% If strMode = "A" Then Response.Write("?PF=" & lngUserProfileID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>" title="<% = strTxtNewUpload %>" class="fileManLink"><% = strTxtNewUpload %></a>&nbsp; 
     &nbsp;<img src="<% = strImagePath %>file_delete.<% = strForumImageType %>" width="16" height="16" alt="<% = strTxtDeleteFile %>" title="<% = strTxtDeleteFile %>" /> <a href="javascript:deleteFile(document.getElementById('fileName').value);" title="<% = strTxtDeleteFile %>" class="fileManLink"><% = strTxtDeleteFile %></a></td>
    </tr>
   </table>
   <br />
   <strong><% = strTxtFilePreview %></strong>
   <br />
   <div align="center">
    <iframe src="RTE_popup_link_preview.asp<% = strQsSID1 %>" id="prevWindow" width="98%" height="263px" style="border: #A5ACB2 1px solid;background-color: #FFFFFF;"></iframe>
   </div>
  </td>
 </tr>
</table>
<br />
<div align="center"><%

	
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnLCode = True Then
	
	If blnTextLinks = True Then 
		Response.Write("<span class=""text"" style=""font-size:10px"">Forum Software by <a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank""  style=""font-size:10px"">Web Wiz Forums&reg;</a> version " & strVersion & "</span>")
	Else
  		Response.Write("<a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank"" ><img src=""webwizforums_image.asp"" border=""0"" title=""Forum Software by Web Wiz Forums&reg; version " & strVersion& """ alt=""Forum Software by Web Wiz Forums&reg; version " & strVersion& """ /></a>")
	End If
	
	Response.Write("<br /><span class=""text"" style=""font-size:10px"">Copyright &copy;2001-2019 Web Wiz Ltd.</span>")
End If 
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******


'Display the process time
If blnShowProcessTime Then Response.Write "<span class=""smText""><br /><br />" & strTxtThisPageWasGeneratedIn & " " & FormatNumber(Timer() - dblStartTime, 3) & " " & strTxtSeconds & "</span>"
%>
</div>
<%

'If an image has been uploaded show it in the preview window
If blnFileUploadDetails Then 
	Response.Write("<script language=""JavaScript"">")
	Response.Write("uploadPreview()")
	Response.Write("</script>")
End If

'Display an alert message letting the user know the topic has been deleted
If Request.QueryString("DL") = "True" Then
	Response.Write("<script language=""JavaScript"">")
	Response.Write("alert('" & strTxtTheFileNowDeleted & "')")
	Response.Write("</script>")
End If

'Display an alert message letting the user know the topic has been deleted
If Request.QueryString("UL") = "true" Then
	Response.Write("<script language=""JavaScript"">")
	Response.Write("alert('" & strTxtYourFileHasBeenSuccessfullyUploaded & "')")
	Response.Write("</script>")
End If
%>
<!-- #include file="includes/footer.asp" -->