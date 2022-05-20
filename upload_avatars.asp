<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="functions/functions_upload.asp" -->
<!--#include file="functions/functions_date_time_format.asp" -->
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
Dim strImageURL
Dim strImageAltText
Dim strAlign
Dim intBorder
Dim lngHorizontal
Dim lngVerical
Dim intWidth
Dim intHeight
Dim strMessageBoxType	'Holds the type of message box used RTE or basic
Dim blnInsertImage
Dim strImageUploadPath	'Holds the path and folder the uploaded files are stored in
Dim saryImageUploadTypes'Holds the array of file to upload


blnInsertImage = false
strImageUploadPath = strUploadFilePath



'Check the user is registered and so able to post
If  bannedIP() OR intGroupID = 2 OR blnActiveMember = False OR blnBanned OR blnAvatarUploadEnabled = False Then

	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If





'Read in the file types that can be uploaded
If blnAvatarUploadEnabled Then

	'Read in the configuration details from the recordset
	strImageTypes = strAvatarTypes
	lngUploadMaxImageSize = intMaxAvatarSize
	
	'If this is a post back then upload the image
	If Request.QueryString("PB") = "Y" Then
	
		'Call upload file function
		strImageName = fileUpload("image")
	
	End If
End If

'Clean up
Call closeDatabase()


'Change \ for /
strImageName = Replace(strImageName, "\", "/", 1, -1, 1)

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>
<% = strTxtAvatarUpload %>
</title>
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

//Function intilise page
function initilzeElements(){
<%

'If an image has been uploaded update the form
If lngErrorFileSize = 0 AND blnExtensionOK = True AND blnFileSpaceExceeded = False AND blnFileExists = False AND strImageName <> "" Then
	
	
	Response.Write(vbCrLf & "	document.getElementById('prevWindow').contentWindow.location.href = '" & Replace(strImageUploadPath, "\", "/", 1, -1, 1)& "/" & strImageName & "'")
	Response.Write(vbCrLf & "	document.getElementById('fileName').value = '" & strImageName & "';")
	Response.Write(vbCrLf & "	document.getElementById('fileNameDisplay').innerHTML = '" & strImageName & "';")

End If
%>
}

//Function to insert avatar in main page
function insertAvatar(){

	if (document.getElementById('fileName').value == ''){
		window.close();
	}else{
		window.opener.document.getElementById('txtAvatar').focus();
		window.opener.document.getElementById('txtAvatar').value = '<% = strUploadFilePath  %>/' + document.getElementById('fileName').value;
		window.opener.document.getElementById('avatar').src = '<% = strUploadFilePath  %>/' + document.getElementById('fileName').value;
		window.close();
	}
}


//Function to check upload file is selected
function checkFile(){
	if (document.getElementById('file').value==''){
	
		alert('<% = strTxtErrorUploadingImage %>\n<% = strTxtNoImageToUpload %>')
		return false;
	}else{
		alert('<% = strTxtPleaseWaitWhileImageIsUploaded %>');
		return true;
	}
}
</script>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
</head>
<body style="margin:0px;" OnLoad="self.focus(); initilzeElements();">
<table width="100%"  border="0" cellpadding="3" cellspacing="0" class="RTEtableTopRow">
 <tr class="RTEtableTopRow">
  <td colspan="2"><h1>
    <% = strTxtAvatarUpload %>
   </h1></td>
 </tr>
 <tr>
  <td colspan="2" class="RTEtableRow"><table width="100%" border="0" align="center" cellpadding="1" cellspacing="0">
    <tr>
     <td width="50%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="2">
       <tr>
        <td width="88%" class="text"><% = strTxtPath %>: <span id="path"><% = strImageUploadPath %></span></td>
       </tr>
       <tr>
        <td class="text"><% = strTxtFileName %>:<br />
         <iframe src="file_browser.asp<% = strQsSID1 %>" id="fileWindow" width="333" height="272px" style="border: #A5ACB2 1px solid;background-color: #FFFFFF;"></iframe></td>
       </tr>
      </table></td>
     <td width="50%" valign="top"><form method="post" action="upload_avatars.asp?PB=Y<% = strQsSID2 %>" name="frmUpload" enctype="multipart/form-data" onsubmit="return checkFile();" >
       <table width="100%" border="0" cellspacing="0" cellpadding="1">
        <tr>
         <td><strong><% = strTxtAvatarUpload %></strong></td>
        </tr>
        <tr>
         <td width="80%"><input id="file" name="file" type="file" size="38" /></td>
        </tr>
        <tr>
         <td><span class="smText"><% Response.Write(strTxtImagesMustBeOfTheType & ", " & Replace(strImageTypes, ";", ", ", 1, -1, 1) & ", " & strTxtAndHaveMaximumFileSizeOf & " " & lngUploadMaxImageSize & "KB")  %></span></td>
        </tr>
        <tr>
         <td><input name="upload" type="submit" id="upload" value="Upload" /></td>
        </tr>
       </table>
      </form>
      <strong><% = strTxtFileProperties %></strong>
     <table width="100%" border="0" cellspacing="1" cellpadding="0">
       <tr>
        <td align="right" width="25%"><% = strTxtFileName %>: </td>
        <td width="75%"><span id="fileNameDisplay"></span> <input type="hidden" name="fileName" id="fileName" /> <span id="fileDownload" style="display:none;"></span>  <span id="fileSize" style="display:none;"></span></td>
       </tr>
       <tr>
        <td align="right"><% = strTxtFileType %>: </td>
        <td><span id="fileType"></span></td>
       </tr>
      </table>
      <strong><% = strTxtFilePreview %></strong> <br />
      <iframe src="RTE_popup_link_preview.asp<% = strQsSID1 %>" id="prevWindow" style="border: #A5ACB2 1px solid;background-color: #FFFFFF;height:123px;width:123px;"></iframe></td>
    </tr>
   </table></td>
 </tr>
 <tr>
  <td width="38%" valign="top" class="RTEtableBottomRow">&nbsp;</td>
  <td width="24%" align="right" class="RTEtableBottomRow">
   <input type="button" name="Submit" id="Submit" value="     <% = strTxtOK %>     " onclick="insertAvatar();">
   &nbsp;
   <input type="button" name="cancel" value=" <% = strTxtCancel %> " onclick="window.close();">
   <br />
   <br />
  </td>
 </tr>
</table>
</body>
</html>
<%
'If the file space is exceeded then tell the user
If blnFileSpaceExceeded Then
	Response.Write("<script language=""JavaScript"">")
	Response.Write("alert('" & strTxtErrorUploadingImage & ".\n" & strTxtAllotedFileSpaceExceeded & " " & intUploadAllocatedSpace & "MB.\n" & strTxtDeleteFileOrImagesUingCP & "');")
	Response.Write("</script>")
	
'If the file already exists tell the user
ElseIf blnFileExists Then 
	Response.Write("<script language=""JavaScript"">")
	Response.Write("alert('" & strTxtErrorUploadingImage & ".\n" & strTxtFileAlreadyUploaded & ".\n" & strTxtSelectOrRenameFile & "');")
	Response.Write("</script>")
	
'If the file upload has failed becuase of the wrong extension display an error message
ElseIf blnExtensionOK = False Then
	Response.Write("<script language=""JavaScript"">")
	Response.Write("alert('" & strTxtErrorUploadingImage & ".\n" &strTxtImageOfTheWrongFileType & ".\n" & strTxtImagesMustBeOfTheType & ", "  &  Replace(strImageTypes, ";", ", ", 1, -1, 1) & "');")
	Response.Write("</script>")

'Else if the file upload has failed becuase the size is to large display an error message
ElseIf lngErrorFileSize <> 0 Then
	Response.Write("<script language=""JavaScript"">")
	Response.Write("alert('" & strTxtErrorUploadingImage & ".\n" & strTxtImageFileSizeToLarge & " " & lngErrorFileSize & "KB.\n" & strTxtMaximumFileSizeMustBe & " " & lngUploadMaxImageSize & "KB');")
	Response.Write("</script>")
End If
%>
