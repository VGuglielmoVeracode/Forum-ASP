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




'Set the timeout of the page
Server.ScriptTimeout =  1000

'Set the response buffer to true as we maybe redirecting
Response.Buffer = True



'Dimension veriables
Dim strMode 		'Holds the page mode (eg admin)
Dim lngUserProfileID	'Holds the profile ID of the user
Dim dblFileSpaceUsed	'Holds the amount of file space used
Dim strFileName		'Name of newly uploaded file
Dim objUploadProgress
Dim strAspUploadPID
Dim strAspUploadBarRef
Dim strMaxImageUpload
Dim strMaxFileUpload
Dim strErrorUploadSize
Dim strXID




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





'Setup for progress bar
If strUploadComponent = "AspUpload" Then
	'Set error trapping
	On Error Resume Next

	'Create AspUpload Progress componnet
	Set objUploadProgress = Server.CreateObject("Persits.UploadProgress")
	strAspUploadPID = "&PID=" & objUploadProgress.CreateProgressID()
	strAspUploadBarRef = "AspUpload_ProgressBar_Frame.asp?to=10" & strAspUploadPID
	Set objUploadProgress = Nothing
	
	'Check to see if an error has occurred
	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then	Call errorMsg("An error has occurred.<br />Please check the Persits AspUpload Component 3.0 or above is installed on the server.", "create_AspUpload_progress_object", "file_upload.asp")
		
	'Disable error trapping
	On Error goto 0
End If





'If this is a postback upload the file
If Request.QueryString("PB") = "Y" Then
	
	'Check the session ID to stop spammers using the email form
	Call checkFormID(Request.QueryString("XID"))
	
	'Call upoload file function
	If Request.QueryString("T") = "image" AND blnImageUpload Then
		strFileName = fileUpload("image")
	ElseIf blnAttachments Then
		strFileName = fileUpload("file")
	End If
	
	'If the file has been uploaded then redirect to file manager
	If lngErrorFileSize = 0 AND blnExtensionOK = True AND blnFileSpaceExceeded = False AND blnFileExists = False AND blnSecurityScanFail = False AND strFileName <> "" Then
	
		'Go back to file manager
		If strMode = "A" Then 
			Response.Redirect("file_manager.asp?UL=true&PF=" & lngUserProfileID & "&M=A&FN=" & strFileName & strQsSID3)
		Else
			Response.Redirect("file_manager.asp?UL=true&FN=" & strFileName & strQsSID3)
		End If
	
	End If
	
	'Calculate the error file upload size in MB
	If lngErrorFileSize >= 1024 Then 
		strErrorUploadSize = FormatNumber((lngErrorFileSize / 1024), 1) & " MB"
	ElseIf lngErrorFileSize > 0 Then 
		strErrorUploadSize = lngErrorFileSize & " KB"
	End If
End If





'If active users is enabled update the active users application array
If blnActiveUsers Then
	'Call active users function
	saryActiveUsers = activeUsers("", strTxtFileManager, "file_manager.asp", 0)
End If


'get session key
strXID = getSessionItem("KEY")


'Clean up
Call closeDatabase()



'Calculate the image upload size in MB
If lngUploadMaxImageSize >= 1024 Then 
	strMaxImageUpload = FormatNumber((lngUploadMaxImageSize / 1024), 1) & " MB"
Else 
	strMaxImageUpload = lngUploadMaxImageSize & " KB"
End If

'Calculate the file upload size in MB
If lngUploadMaxFileSize >= 1024 Then 
	strMaxFileUpload = FormatNumber((lngUploadMaxFileSize / 1024), 1) & " MB"
Else 
	strMaxFileUpload = lngUploadMaxFileSize & " KB"
End If



'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & "<a href=""file_manager.asp"
If strMode = "A" Then 
	strBreadCrumbTrail = strBreadCrumbTrail & "?PF=" & lngUserProfileID & "&M=A" & strQsSID2 
Else 
	strBreadCrumbTrail = strBreadCrumbTrail & strQsSID1
End If
strBreadCrumbTrail = strBreadCrumbTrail & """>" & strTxtFileManager & "</a>" & strNavSpacer & strTxtFileUpload


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
//Function swap link type
function swapLinkType(selType){
	if (selType.value == "fileUp"){
		document.getElementById("imageProperties").style.display="none";
    		document.getElementById("fileProperties").style.display="block";<%
    		
'If this is Gekco based browser or Opera the element needs to be set to visable
'If RTEenabled = "Gecko" OR RTEenabled = "opera" Then Response.Write(vbCrLf & "		document.getElementById(""fileProperties"").style.visibility=""visable"";") 		
    		%>
    		
	}else{
		document.getElementById("fileProperties").style.display="none";
		document.getElementById("imageProperties").style.display="block";<%
		
'If this is Gekco based browser or Opera the element needs to be set to visable
'If RTEenabled = "Gecko" OR RTEenabled = "opera" Then Response.Write(vbCrLf & "		document.getElementById(""imageProperties"").style.visibility=""visable"";") 		
    		%>
	}
}

//Function to check upload file is selected
function checkFile(){<%

'Select which element to check for a file/image upload	
If blnAttachments AND blnImageUpload Then
%>
	if ((document.getElementById('file1').value=='')&&(document.getElementById('file2').value=='')) {<%
ElseIf blnAttachments Then
%>
	if (document.getElementById('file1').value=='') {<%	
ElseIf blnImageUpload Then
%>
	if (document.getElementById('file2').value=='') {<%
End If
%>	
		alert('<% = strTxtErrorUploadingFile %>\n<% = strTxtNoFileToUpload %>')
		return false;
	}else{<%
		
'AspUpload Progress bar
If strUploadComponent = "AspUpload" Then

%>
		winOpener('<% = strAspUploadBarRef %>', 'progressBar', 0, 0, 410, 190);<%

Else
%>
		alert('<% = strTxtPleaseWaitWhileFileIsUploaded %>');<%
End If

%>
		return true;
	}
}

</script>

<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />  	
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% = strTxtFileManager & " - " & strTxtFileUpload  %></h1></td>
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
  <td valign="top"><%

'If both image and file uploads are permitted for this member give then the option
If blnAttachments AND blnImageUpload Then

%>
    <% = strTxtSelectUploadType %>
   <select name="selType" id="selType" onchange="swapLinkType(this)">
    <option value="imageUp"><% = strTxtImageUpload %></option>
    <option value="fileUp"<% If Request.QueryString("T") = "file" Then Response.Write(" selected") %>><% = strTxtFileUpload %></option>
   </select> <%

End If


'If attachments allowed 
If blnAttachments Then
%>    
   <div id="fileProperties"<% If Request.QueryString("T") <> "file" AND blnImageUpload Then Response.Write(" style=""display:none""") %>>          
    <br />
    <strong><% = strTxtFileUpload %></strong>
    <br />
    <form method="post" action="file_upload.asp?PB=Y&T=file<% If strMode = "A" Then Response.Write("&PF=" & lngUserProfileID & "&M=A") %><% = strAspUploadPID %>&XID=<% = strXID & strQsSID2 %>" name="frmUpload" enctype="multipart/form-data" onsubmit="return checkFile();" >
     <input id="file1" name="file" type="file" size="50" />
     <br />
     <br />
     <% Response.Write(strTxtFilesMustBeOfTheType & ", " & Replace(strUploadFileTypes, ";", ", ", 1, -1, 1) & ", " & strTxtAndHaveMaximumFileSizeOf & " " & strMaxFileUpload) %>
     <br />
     <br />
     <input name="upload" type="submit" id="upload1" value="<% = strTxtUpload & " " & strTxtFile %>">
     <br />
     <br />
    </form>
   </div><%

End If


'If image upload allowed
If blnImageUpload Then

%>
   <div id="imageProperties"<% If Request.QueryString("T") = "file" Then Response.Write(" style=""display:none""") %>> 
    <br />
    <strong><% = strTxtImageUpload %></strong>
    <br />
    <form method="post" action="file_upload.asp?PB=Y&T=image<% If strMode = "A" Then Response.Write("&PF=" & lngUserProfileID & "&M=A") %><% = strAspUploadPID %>&XID=<% = strXID & strQsSID2 %>" name="frmUpload" enctype="multipart/form-data" onsubmit="return checkFile();" >
     <input id="file2" name="file" type="file" size="50" />
     <br />
     <br />
     <% Response.Write(strTxtImagesMustBeOfTheType & ", " & Replace(strImageTypes, ";", ", ", 1, -1, 1) & ", " & strTxtAndHaveMaximumFileSizeOf & " " & strMaxImageUpload)  %>
     <br />
     <br />
     <input name="upload" type="submit" id="upload2" value="<% = strTxtUpload & " " & strTxtImage %>">
     <br />
     <br />
    </form>
   </div>
   <%

End If

%>
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

'If the file space is exceeded then tell the user
If blnFileSpaceExceeded Then
	Response.Write("<script language=""JavaScript"">")
	Response.Write("alert('" & strTxtErrorUploadingFile & ".\n" & strTxtAllotedFileSpaceExceeded & " " & intUploadAllocatedSpace & "MB.\n" & strTxtDeleteFileOrImagesUingCP & "');")
	Response.Write("</script>")
	
'If the file already exists tell the user
ElseIf blnFileExists Then 
	Response.Write("<script language=""JavaScript"">")
	Response.Write("alert('" & strTxtErrorUploadingFile & ".\n" & strTxtFileAlreadyUploaded & ".\n" & strTxtSelectOrRenameFile & "');")
	Response.Write("</script>")
	
'If the file upload has failed becuase of the wrong extension display an error message
ElseIf blnExtensionOK = False Then
	Response.Write("<script language=""JavaScript"">")
	Response.Write("alert('" & strTxtErrorUploadingFile & ".\n" &strTxtFileOfTheWrongFileType & ".\n" & strTxtFilesMustBeOfTheType & ", "  &  Replace(strUploadFileTypes, ";", ", ", 1, -1, 1) & "');")
	Response.Write("</script>")

'Else if the file upload has failed becuase the size is to large display an error message
ElseIf lngErrorFileSize <> 0 AND Request.QueryString("T") = "image" Then
	Response.Write("<script language=""JavaScript"">")
	Response.Write("alert('" & strTxtErrorUploadingImage & ".\n" & strTxtImageFileSizeToLarge & " " & strErrorUploadSize & ".\n" & strTxtMaximumFileSizeMustBe & " " & strMaxImageUpload & "');")
	Response.Write("</script>")
	
'Else if the file upload has failed becuase the size is to large display an error message
ElseIf lngErrorFileSize <> 0 Then
	Response.Write("<script language=""JavaScript"">")
	Response.Write("alert('" & strTxtErrorUploadingFile & ".\n" & strTxtFileSizeToLarge & " " & strErrorUploadSize & ".\n" & strTxtMaximumFileSizeMustBe & " " & strMaxFileUpload & "');")
	Response.Write("</script>")

'Else if the security scan failed
ElseIf blnSecurityScanFail Then
	Response.Write("<script language=""JavaScript"">")
	Response.Write("alert('" & strTxtErrorUploadingFile & ".\n" & strTxtTheFileFailedTheSecurityuScanAndHasBeenDeleted & "');")
	Response.Write("</script>")

End If




%>
<!-- #include file="includes/footer.asp" -->