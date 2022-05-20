<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="RTE_configuration/RTE_setup.asp" -->
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



'Intiliase variables
blnExtensionOK = True
strUploadFilePath = strUploadFilePath


'read in the forum ID and message box type
intForumID = IntC(getSessionItem("FID"))


'Check the user is welcome in this forum
Call forumPermissions(intForumID, intGroupID)


'If the user is user is using a banned IP redirect to an error page
If bannedIP() OR blnBanned OR blnImageUpload = false Then
	
	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If




'Read in the file types that can be uploaded
If blnImageUpload Then
	
	'If this is a post back then upload the image
	If Request.QueryString("PB") = "Y" Then
	
	
		'Call upoload file function
		strImageName = fileUpload("image")
	
	End If
End If







'Reset Server Objects
Call closeDatabase()

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title>Image Upload</title>

<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write(vbCrLf & "<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Rich Text Editor(TM) ver. " & strRTEversion & "" & _
vbCrLf & "Info: http://www.richtexteditor.org" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>

<script  language="JavaScript">
<%
'If the image has been saved then place it in the post
If lngErrorFileSize = 0 AND blnExtensionOK = True AND blnFileSpaceExceeded = False AND blnFileExists = False AND blnSecurityScanFail = False AND strImageName <> "" Then
%>
	window.opener.document.getElementById('message').focus();
	window.opener.document.getElementById('message').value += '[IMG]<% = strUploadFilePath & "/" & strImageName %>[/IMG]';
	window.opener.document.getElementById('uploads').value += '<% = strImageName %>;';
	window.close();
<%
End If
%>
</script>
<style type="text/css">
<!--
html, body {
  background: ButtonFace;
  color: ButtonText;
  font: font-family: Verdana, Arial, Helvetica, sans-serif;
  font-size: 12px;
  margin: 2px;
  padding: 4px;
}
.text {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 12px;
	color: #000000;
}
.error {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 12px;
	color: #FF0000;
}
legend {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	color: #0000FF;
}
-->
</style>
</head>
<body OnLoad="self.focus(); document.forms.frmImageUp.Submit.disabled=true;"><%

'If the user is allowed to upload then show them the form
If blnImageUpload Then	

	%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="0">
 <form action="non_RTE_upload_images.asp?PB=Y<% = strQsSID2 %>" method="post" enctype="multipart/form-data" name="frmImageUp" id="frmImageUp" onSubmit="alert('<% = strTxtPleaseWaitWhileImageIsUploaded %>')">
  <tr> 
   <td width="100%"> 
    <fieldset>
    <legend><% = strTxtImageUpload %></legend>
    <table width="100%" border="0" cellspacing="0" cellpadding="1">
     <tr> 
      <td width="10%" align="right" class="text"><% = strTxtImage %>:</td>
      <td width="90%"><input name="file" name="id" type="file" size="35" onFocus="document.forms.frmImageUp.Submit.disabled=false;" onclick="document.forms.frmImageUp.Submit.disabled=false;" />
        </td>
     </tr>
     <tr align="center"> 
      <td colspan="2" class="text"><br /><% 
      	
      	'If the file space is exceeded then tell the user
	If blnFileSpaceExceeded Then
		
		Response.Write("<span class=""error"">" & strTxtAllotedFileSpaceExceeded & " " & intUploadAllocatedSpace & "MB. " & strTxtDeleteFileOrImagesUingCP & "</span>")
		
	'If the file already exists tell the user
	ElseIf blnFileExists Then 
		
		Response.Write("<span class=""error"">" & strTxtFileAlreadyUploaded & ". " & strTxtSelectOrRenameFile & "</span>")
		
	'If the file upload has failed becuase of the wrong extension display an error message
	ElseIf blnExtensionOK = False Then

		Response.Write("<span class=""error"">" & strTxtImageOfTheWrongFileType & ".<br />" & strTxtImagesMustBeOfTheType & ", "  &  Replace(strImageTypes, ";", ", ", 1, -1, 1) & "</span>")

	'Else if the file upload has failed becuase the size is to large display an error message
	ElseIf lngErrorFileSize <> 0 Then

		Response.Write("<span class=""error"">" & strTxtImageFileSizeToLarge & " " & lngErrorFileSize & "KB.<br />" & strTxtMaximumFileSizeMustBe & " " & lngUploadMaxImageSize & "KB</span>")
	
	'Else if the security scan failed
	ElseIf blnSecurityScanFail Then
		
		Response.Write("<span class=""error"">" & strTxtErrorUploadingFile & ".<br />" & strTxtTheFileFailedTHeSecurityuScanAndHasBeenDeleted & "</span>")
		
	'Else display a message of the image types that can be uploaded
	Else
      
      		Response.Write(strTxtImagesMustBeOfTheType & ", " & Replace(strImageTypes, ";", ", ", 1, -1, 1) & ", " & strTxtAndHaveMaximumFileSizeOf & " " & lngUploadMaxImageSize & "KB") 
      	
      	End If
      %></td>
     </tr>
    </table>
    </fieldset></td>
  </tr>
  <tr align="right"> 
   <td> <input type="submit" name="Submit" value="     <% = strTxtOK %>     "> &nbsp; <input type="button" name="cancel" value=" <% = strTxtCancel %> " onclick="window.close()"></td>
  </tr>
 </form>
</table><%

End If

%>
</body>
</html>