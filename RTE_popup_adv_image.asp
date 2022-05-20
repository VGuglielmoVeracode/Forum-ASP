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
'**  Web Wiz Rich Text Editor(TM)
'**  http://www.richtexteditor.org
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


'If the page is opening in SSL then keep it running in SSL
If Request.ServerVariables("HTTPS") = "on" Then blnSslEnabledPage = True



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
Dim objUploadProgress
Dim strAspUploadPID
Dim strAspUploadBarRef
Dim strMaxImageUpload
Dim strErrorUploadSize


blnInsertImage = false
strImageUploadPath = strUploadFilePath




	
	
'If the user is user is using a banned IP redirect to an error page
If bannedIP() OR blnBanned Then
		
	'Clean up
	Call closeDatabase()
		
	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If




'Setup for progress bar
If strUploadComponent = "AspUpload"  AND blnImageUpload Then
	
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




'If this is a post back then upload the image (use querysting as it is a multipart/form-data form)
If Request.QueryString("PB") = "Y" AND blnImageUpload Then

	'Call upoload file function
	strImageName = fileUpload("image")
	
	'Calculate the error file upload size in MB
	If lngErrorFileSize >= 1024 Then 
		strErrorUploadSize = FormatNumber((lngErrorFileSize / 1024), 1) & " MB"
	ElseIf lngErrorFileSize > 0 Then 
		strErrorUploadSize = lngErrorFileSize & " KB"
	End If



'If this a normal form post back to insert an image read in the form elements
ElseIf Request.Form("URL") <> "http://" AND Request.Form("URL") <> "" AND Request.Form("postBack") Then
	
	'Initilise variable
	intBorder = 0
	
	'Get form elements
	strImageURL = Request.Form("URL")
	strImageAltText = Request.Form("Alt")
	strAlign = Request.Form("align")
	If isNumeric(Request.Form("intBorder")) Then intBorder = Request.Form("border")
	If isNumeric(Request.Form("hoz")) Then lngHorizontal = LngC(Request.Form("hoz"))
	If isNumeric(Request.Form("vert")) Then lngVerical = LngC(Request.Form("vert"))
	If isNumeric(Request.Form("width")) Then intWidth = LngC(Request.Form("width"))
	If isNumeric(Request.Form("height")) Then intHeight = LngC(Request.Form("height"))
	
	'Escape characters that will course a crash
	strImageURL = Replace(strImageURL, "'", "\'", 1, -1, 1)
	strImageURL = Replace(strImageURL, """", "\""", 1, -1, 1)
	strImageAltText = Replace(strImageAltText, "'", "\'", 1, -1, 1)
	strImageAltText = Replace(strImageAltText, """", "\""", 1, -1, 1)
	
	blnInsertImage = true
End If


'Clean up
Call closeDatabase()


'Calculate the image upload size in MB
If lngUploadMaxImageSize >= 1024 Then 
	strMaxImageUpload = FormatNumber((lngUploadMaxImageSize / 1024), 1) & " MB"
Else 
	strMaxImageUpload = lngUploadMaxImageSize & " KB"
End If


'Change \ for /
strImageName = Replace(strImageName, "\", "/", 1, -1, 1)

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Image Properties</title>

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


'If this is Gecko based browser or Opera link to JS code for Gecko
If RTEenabled = "Gecko" OR RTEenabled = "opera" Then Response.Write(vbCrLf & "<script language=""JavaScript"" src=""RTE_javascript_gecko.asp"" type=""text/javascript""></script>")
	
%>
<script language="JavaScript">

//Function intilise page
function initilzeElements(){
<%

'If an image has been uploaded update the form
If lngErrorFileSize = 0 AND blnExtensionOK = True AND blnFileSpaceExceeded = False AND blnFileExists = False AND blnSecurityScanFail = False AND strImageName <> "" Then
	
	'See if we are to use the full URL for file
	If blnUseFullURLpath Then
		Response.Write(vbCrLf & "	document.getElementById('URL').value = '" & strFullURLpathToRTEfiles & Replace(strImageUploadPath, "\", "/", 1, -1, 1)  & "/" & strImageName & "'")
		Response.Write(vbCrLf & "	document.getElementById('prevWindow').contentWindow.document.getElementById('prevFile').src = '" & strFullURLpathToRTEfiles & Replace(strImageUploadPath, "\", "/", 1, -1, 1)  & "/" & strImageName & "'")
	Else
		Response.Write(vbCrLf & "	document.getElementById('URL').value = '" & Replace(strImageUploadPath, "\", "/", 1, -1, 1)  & "/" & strImageName & "'")
		Response.Write(vbCrLf & "	document.getElementById('prevWindow').contentWindow.document.getElementById('prevFile').src = '" & Replace(strImageUploadPath, "\", "/", 1, -1, 1)  & "/" & strImageName & "'")
	End If
	
	Response.Write(vbCrLf & "	document.getElementById('Submit').disabled = false;")

'Else no image has been uploaded so just initilise the form
Else
	Response.Write(vbCrLf & "	document.getElementById('URL').value = 'http://'")
	Response.Write(vbCrLf & "	document.getElementById('Submit').disabled = true;")
End If
%>
}

<%
'If this a post back write javascript
If blnInsertImage Then
	
	Response.Write(vbCrLf & vbCrLf & "	editor = window.opener.document.getElementById('WebWizRTE');")
	
	'Tell that we are an image
	Response.Write(vbCrLf & vbCrLf & "	img = editor.contentWindow.document.createElement('img');")
	
	'Set image attributes
	Response.Write(vbCrLf & vbCrLf & "	img.setAttribute('src', '" & strImageURL & "');")
	Response.Write(vbCrLf & "	img.setAttribute('border', '" & intBorder & "');")
	If strImageAltText <> "" Then Response.Write(vbCrLf & "	img.setAttribute('alt', '" & strImageAltText & "');")
	If lngHorizontal <> "" Then Response.Write(vbCrLf & "	img.setAttribute('hspace', '" & lngHorizontal & "');")
	If intWidth <> "" Then Response.Write(vbCrLf & "	img.setAttribute('width', '" & intWidth & "');")
	If intHeight <> "" Then Response.Write(vbCrLf & "	img.setAttribute('height', '" & intHeight & "');")
	If lngVerical <> "" Then Response.Write(vbCrLf & "	img.setAttribute('vspace', '" & lngVerical & "');")
	If strAlign <> "" Then Response.Write(vbCrLf & "	img.setAttribute('align', '" & strAlign & "');")
	 
     
     	'If this is Mozilla or Opera then we need to call insertElementPosition to find where to place the image
     	If RTEenabled = "Gecko" OR RTEenabled = "opera" Then 
		
		Response.Write(vbCrLf & vbCrLf & "	try{" & _
					vbCrLf & "		insertElementPosition(editor.contentWindow, img);" & _
					vbCrLf & "	}catch(exception){" & _
					vbCrLf & "		alert('" & strTxtErrorInsertingObject & "');" & _
					vbCrLf & "		editor.contentWindow.focus();" & _
					vbCrLf & "	}")
	
	'Else this is IE so placing the image is simpler
	Else
		Response.Write(vbCrLf & vbCrLf & "	try{" & _
					vbCrLf & "		editor.contentWindow.focus();" & _
					vbCrLf & "		editor.contentWindow.document.selection.createRange().pasteHTML(img.outerHTML);" & _
					vbCrLf & "	}catch(exception){" & _
					vbCrLf & "		alert('" & strTxtErrorInsertingObject & "');" & _
					vbCrLf & "		editor.contentWindow.focus();" & _
					vbCrLf & "	}")
	End If
		
	'Set focus
	'If Opera change the focus method
	If RTEenabled = "opera" Then
		
		Response.Write(vbCrLf & "	editor.focus();")
	Else
		Response.Write(vbCrLf & "	editor.contentWindow.focus();")
	End If
		
	'Close window
	Response.Write(vbCrLf & "	window.close();")
End If

%>

//Function to preview image
function getImage(URL){
	document.getElementById("prevWindow").contentWindow.document.getElementById("prevFile").src = URL
}

//Function to change image properties
function changeImage(){
	document.getElementById("prevWindow").contentWindow.document.getElementById("prevFile").alt = document.getElementById('Alt').value
	document.getElementById("prevWindow").contentWindow.document.getElementById("prevFile").align = document.getElementById('align').value
	document.getElementById("prevWindow").contentWindow.document.getElementById("prevFile").border = document.getElementById('border').value
	document.getElementById("prevWindow").contentWindow.document.getElementById("prevFile").hspace = document.getElementById('hoz').value
	document.getElementById("prevWindow").contentWindow.document.getElementById("prevFile").vspace = document.getElementById('vert').value
	//Check a value for width and hieght is set or image will be deleted
	if (document.getElementById('width').value!=''){
		document.getElementById("prevWindow").contentWindow.document.getElementById("prevFile").width = document.getElementById('width').value
	}
	if (document.getElementById('height').value!=''){
		document.getElementById("prevWindow").contentWindow.document.getElementById("prevFile").height = document.getElementById('height').value
	}
}

<%
'If image upload is enabled then have the following function
If blnImageUpload Then	
%>
//Function to check upload file is selected
function checkFile(){
	if (document.getElementById('file').value==''){
	
		alert('<% = strTxtErrorUploadingImage %>\n<% = strTxtNoImageToUpload %>')
		return false;
	}else{<%
		
'AspUpload Progress bar
If strUploadComponent = "AspUpload" Then

%>
		winOpener('<% = strAspUploadBarRef %>', 'progressBar', 0, 0, 410, 190);<%

Else
%>
		alert('<% = strTxtPleaseWaitWhileImageIsUploaded %>');<%
End If

%>
		return true;
	}
}<%
End If
%>

//function to open pop up window
function winOpener(theURL, winName, scrollbars, resizable, width, height) {

	winFeatures = 'left=' + (screen.availWidth-10-width)/2 + ',top=' + (screen.availHeight-30-height)/2 + ',scrollbars=' + scrollbars + ',resizable=' + resizable + ',width=' + width + ',height=' + height + ',toolbar=0,location=0,status=1,menubar=0'
  	window.open(theURL, winName, winFeatures);
}
</script>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
</head>
<body style="margin:0px;" OnLoad="self.focus(); initilzeElements();">
<table width="100%"  border="0" cellpadding="3" cellspacing="0" class="RTEtableTopRow">
    <tr class="RTEtableTopRow">
      <td colspan="2"><h1><% = strTxtImageProperties %></h1></td>
    </tr>
    <tr>
      <td colspan="2" class="RTEtableRow"><table width="100%" border="0" align="center" cellpadding="1" cellspacing="0">
        <tr>
          <td width="38%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="2">
            <tr>
              <td width="88%" class="text"><% = strTxtPath %>: <span id="path"><% = strImageUploadPath %></span></td>
            </tr>
            <%
            
'If image upload is enabled then display an image upload form
If blnImageUpload Then

%>
            <tr>
              <td class="text"><% = strTxtFileName %>:<iframe src="RTE_file_browser.asp?look=img<% = strQsSID2 %>" id="fileWindow" width="98%" height="180px" style="border: #A5ACB2 1px solid;background-color: #FFFFFF;"></iframe></td>
            </tr>
            <form method="post" action="RTE_popup_adv_image.asp?PB=Y<% = strAspUploadPID %><% = strQsSID2 %>" name="frmUpload" enctype="multipart/form-data" onsubmit="return checkFile();" >
             <tr>
              <td class="text"><% = strTxtImageUpload %></td>
            </tr>
            <tr>
              <td class="smText"><% Response.Write(strTxtImagesMustBeOfTheType & ", " & Replace(strImageTypes, ";", ", ", 1, -1, 1) & ", " & strTxtAndHaveMaximumFileSizeOf & " " & strMaxImageUpload)  %></td>
            </tr>
            <tr>
              <td><input id="file" name="file" type="file" size="35" /></td>
            </tr>
            <tr>
              <td>
              	<input name="upload" type="submit" id="upload" value="Upload">
              </td>
            </tr>
           </form><%
'Else file uploading is disabled so show a larger file browser window
Else

%>
	    <tr>
              <td class="text"><% = strTxtFileName %>:<iframe src="RTE_file_browser.asp?look=img<% = strQsSID2 %>" id="fileWindow" width="98%" height="278px" style="border: #A5ACB2 1px solid;background-color: #FFFFFF;"></iframe></td>
            </tr>
<%
End If

%>
          </table></td>
          <td width="58%" valign="top">
          <form method="post" action="RTE_popup_adv_image.asp<% = strQsSID1 %>" name="frmImageInsrt">
            <table width="100%" border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td width="25%" align="right" class="text"><% = strTxtImageURL %>:</td>
                <td colspan="5"><input name="URL" type="text" id="URL" value="" size="35" onchange="document.getElementById('Submit').disabled=false;" onkeypress="document.getElementById('Submit').disabled=false;">
                  <input name="preview" type="button" id="preview" value="<% = strTxtPreview %>" onclick="getImage(document.getElementById('URL').value)">
                </td>
              </tr>
              <tr>
                <td align="right" class="text"><% = strTxtAlternativeText %>:</td>
                <td colspan="5"><input name="Alt" type="text" id="Alt" size="35" onBlur="changeImage()"></td>
              </tr>
              <tr>
                <td align="right" class="text"><% = strTxtWidth %>:</td>
                <td width="4%"><input name="width" type="text" id="width" size="3" maxlength="3" onkeyup="changeImage()" autocomplete="off" /></td>
                <td width="22%" align="right" class="text"><% = strTxtHorizontal %>:</td>
                <td width="6%"><input name="hoz" type="text" id="hoz" size="3" maxlength="3" onkeyup="changeImage()" autocomplete="off" /></td>
                <td width="20%" align="right" class="text"><% = strTxtAlignment %>:</td>
                <td width="30%"><select size="1" name="align" id="align" onchange="changeImage()">
                  <option value="" selected >Default</option>
                  <option value="left">Left</option>
                  <option value="right">Right</option>
                  <option value="texttop">Texttop</option>
                  <option value="absmiddle">Absmiddle</option>
                  <option value="baseline">Baseline</option>
                  <option value="absbottom">Absbottom</option>
                  <option value="bottom">Bottom</option>
                  <option value="middle">Middle</option>
                  <option value="top">Top</option>
                </select></td>
              </tr>
              <tr>
                <td align="right" class="text"><% = strTxtHeight %>:</td>
                <td><input name="height" type="text" id="height" size="3" maxlength="3" onkeyup="changeImage()" autocomplete="off" /></td>
                <td align="right" class="text"><% = strTxtVertical %>:</td>
                <td><input name="vert" type="text" id="vert" size="3" maxlength="3" onkeyup="changeImage()" autocomplete="off" /></td>
                <td align="right" class="text"><% = strTxtBorder %>:</td>
                <td><input name="border" type="text" id="border" size="3" maxlength="2" onKeyUp="changeImage()" autocomplete="off" /></td>
              </tr>
              <tr>
                <td align="right" valign="top" class="text"><% = strTxtPreview %>:</td>
                <td colspan="5"><iframe src="RTE_popup_image_preview.asp<% = strQsSID1 %>" id="prevWindow" width="98%" height="215px" style="border: #A5ACB2 1px solid;background-color: #FFFFFF;"></iframe></td>
              </tr>
           </table>
         </td>
        </tr>
      </table>
     </td>
    </tr>
    <tr>
      <td width="38%" valign="top" class="RTEtableBottomRow"><%

'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnLCode Then
	Response.Write("<span class=""text"" style=""font-size:10px""><a href=""http://www.richtexteditor.org"" target=""_blank"" style=""font-size:10px"">Web Wiz Rich Text Editor</a> version " & strRTEversion & "</span>")
End If 
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******      
      
      %></td>
      <td width="24%" align="right" class="RTEtableBottomRow">        <input type="hidden" name="postBack" value="true">
        <input type="submit" name="Submit" id="Submit" value="     <% = strTxtOK %>     ">&nbsp;<input type="button" name="cancel" value=" <% = strTxtCancel %> " onclick="window.close()">
        <br /><br />
      </td>
    </tr>
   </form>
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
	Response.Write("alert('" & strTxtErrorUploadingImage & ".\n" & strTxtImageFileSizeToLarge & " " & strErrorUploadSize & ".\n" & strTxtMaximumFileSizeMustBe & " " & strMaxImageUpload & "');")
	Response.Write("</script>")

'Else if the security scan failed
ElseIf blnSecurityScanFail Then
	Response.Write("<script language=""JavaScript"">")
	Response.Write("alert('" & strTxtErrorUploadingFile & ".\n" & strTxtTheFileFailedTheSecurityuScanAndHasBeenDeleted & "');")
	Response.Write("</script>")
End If
%>