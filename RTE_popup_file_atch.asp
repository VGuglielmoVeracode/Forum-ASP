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
Dim strHyperlink
Dim strTitle
Dim strWindow
Dim strErrorMessage	'Holds the error emssage if the file is not uploaded
Dim strMessageBoxType	'Holds the type of message box used RTE or basic
Dim strFileName		'Holds the file name
Dim blnAttachFile
Dim strAttachFilePath
Dim objUploadProgress
Dim strAspUploadPID
Dim strAspUploadBarRef
Dim strMaxFileUpload
Dim strErrorUploadSize



'Intiliase variables
blnAttachFile = false
strAttachFilePath = strUploadFilePath




'If the user is user is using a baned IP redirect to an error page
If bannedIP() OR blnBanned OR blnAttachments = false Then
	
	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If



'Setup for progress bar
If strUploadComponent = "AspUpload"  AND blnAttachments Then
	
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




'If this is a post back then upload the file (use querysting as it is a multipart/form-data form)
If Request.QueryString("PB") = "Y" AND blnAttachments Then
	
	
	'Call upoload file function
	strFileName = fileUpload("file")

	'Calculate the error file upload size in MB
	If lngErrorFileSize >= 1024 Then 
		strErrorUploadSize = FormatNumber((lngErrorFileSize / 1024), 1) & " MB"
	ElseIf lngErrorFileSize > 0 Then 
		strErrorUploadSize = lngErrorFileSize & " KB"
	End If
	


'If this a normal form post back to insert attach the file and read in the form elements
ElseIf Request.Form("URL") <> "http://" AND Request.Form("URL") <> "" AND Request.Form("postBack") Then
	
	'Get form elements
	strHyperlink = Request.Form("URL")
	strTitle = Request.Form("Title")
	strWindow = Request.Form("Window")
	
	'Escape characters that will course a crash
	strHyperlink = Replace(strHyperlink, "'", "\'", 1, -1, 1)
	strHyperlink = Replace(strHyperlink, """", "\""", 1, -1, 1)
	strTitle = Replace(strTitle, "'", "\'", 1, -1, 1)
	strTitle = Replace(strTitle, """", "\""", 1, -1, 1)
	strWindow = Replace(strWindow, "'", "\'", 1, -1, 1)
	strWindow = Replace(strWindow, """", "\""", 1, -1, 1)	
	
	blnAttachFile = true
End If


'Clean up
Call closeDatabase()

'Calculate the file upload size in MB
If lngUploadMaxFileSize >= 1024 Then 
	strMaxFileUpload = FormatNumber((lngUploadMaxFileSize / 1024), 1) & " MB"
Else 
	strMaxFileUpload = lngUploadMaxFileSize & " KB"
End If


'Change \ for /
strAttachFilePath = Replace(strAttachFilePath, "\", "/", 1, -1, 1)
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Attach File Properties</title>

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

'If a file has been uploaded update the form
If lngErrorFileSize = 0 AND blnExtensionOK = True AND blnFileSpaceExceeded = False AND blnFileExists = False AND blnSecurityScanFail = False AND strFileName <> "" Then
	
	'See if we are to use the full URL for file
	If blnUseFullURLpath Then
		Response.Write(vbCrLf & "	document.getElementById('URL').value = '" & strFullURLpathToRTEfiles & Replace(strAttachFilePath, "\", "/", 1, -1, 1) & "/" & strFileName & "'")
		Response.Write(vbCrLf & "	showPreview('" & strFullURLpathToRTEfiles & Replace(strAttachFilePath, "\", "/", 1, -1, 1) & "/" & strFileName & "');")
	Else
		Response.Write(vbCrLf & "	document.getElementById('URL').value = '" & Replace(strAttachFilePath, "\", "/", 1, -1, 1) & "/" & strFileName & "'")
		Response.Write(vbCrLf & "	showPreview('" & Replace(strAttachFilePath, "\", "/", 1, -1, 1) & "/" & strFileName & "');")
	End If
	
	Response.Write(vbCrLf & "	document.getElementById('Submit').disabled = false;")

'Else no file has been uploaded so just initilise the form
Else
	Response.Write(vbCrLf & "	document.getElementById('URL').value = 'http://'")
	Response.Write(vbCrLf & "	document.getElementById('Submit').disabled = true;")
End If
%>
}

<%
'If this a post back write javascript
If blnAttachFile Then
	
	'*********************************************
	'***  	JavaScript for Mozilla & IE	 *****
	'*********************************************
	
	Response.Write(vbCrLf & "	editor = window.opener.document.getElementById('WebWizRTE');")
	
	'Mozilla and Opera use different methods than IE to get the selected text (if any)
	If RTEenabled = "Gecko" OR RTEenabled = "opera" Then 
		Response.Write(vbCrLf & vbCrLf & "	var selectedRange = editor.contentWindow.window.getSelection();")
	Else	
		Response.Write(vbCrLf & vbCrLf & "	var selectedRange = editor.contentWindow.document.selection.createRange();")
	End If	
	
	
	
	
	'If there is a selected area, turn it into a hyperlink
	Response.Write(vbCrLf & vbCrLf & "if (selectedRange != null && selectedRange")
	If RTEenabled = "winIE" Then Response.Write(".text")
	Response.Write(" != ''){")

	'Create hyperlink
	Response.Write(vbCrLf & "	editor.contentWindow.window.document.execCommand('CreateLink', false, '" & strHyperlink & "')")
		
	'Set attributes if required
	If strTitle <> "" OR strWindow <> "" Then
		
		'Set hyperlink attributes
		Response.Write(vbCrLf & vbCrLf & "	var hyperlink = editor.contentWindow.window.document.getElementsByTagName('a');" & _
			       vbCrLf & "	for (var i=0; i < hyperlink.length; i++){" & _
			       vbCrLf & "		if (hyperlink[i].getAttribute('href').search('" & Replace(strHyperlink, "?", "\\?", 1, -1, 1) & "') != -1){")
		
		'Set title, or window if required	       
		If strTitle <> "" Then Response.Write(vbCrLf & "			hyperlink[i].setAttribute('title','" & strTitle & "');")
		If strWindow <> "" Then Response.Write(vbCrLf & "			hyperlink[i].setAttribute('target','" & strWindow & "');")
			       
		Response.Write(vbCrLf & "		}" & _
			       vbCrLf & "	}")
	End If
	
	
	
	'Else no selected area so use the hyperlink text as the displayed text
	Response.Write(vbCrLf & "}else{")
	
	
	'Tell that we are making a hyperlink 'a'
	Response.Write(vbCrLf & vbCrLf & "	link = editor.contentWindow.document.createElement('a');")
	
	
	Response.Write(vbCrLf & vbCrLf & "	link.setAttribute('href', '" & strHyperlink & "');")
	If strTitle <> "" Then Response.Write(vbCrLf & "	link.setAttribute('title', '" & strTitle & "');")
	If strWindow <> "" Then Response.Write(vbCrLf & "	link.setAttribute('target', '" & strWindow & "');")
	
	'Use the selected text range make this the displayed text otherwise use the link as the displayed text
	Response.Write(vbCrLf & vbCrLf & "	link.appendChild(editor.contentWindow.document.createTextNode('" & strHyperlink & "'));")
	
     	'If this is Mozilla then we need to call insertElementPosition to find where to place the file
     	If RTEenabled = "Gecko" OR RTEenabled = "opera" Then 
		
		Response.Write(vbCrLf & vbCrLf & "	try{" & _
					vbCrLf & "		insertElementPosition(editor.contentWindow, link);" & _
					vbCrLf & "	}catch(exception){" & _
					vbCrLf & "		alert('" & strTxtErrorInsertingObject & "');" & _
					vbCrLf & "		editor.contentWindow.focus();" & _
					vbCrLf & "	}")
	
	'Else this is IE so placing the file is simpler
	Else
		Response.Write(vbCrLf & vbCrLf & "	try{" & _
					vbCrLf & "		editor.contentWindow.focus();" & _
					vbCrLf & "		editor.contentWindow.document.selection.createRange().pasteHTML(link.outerHTML);" & _
					vbCrLf & "	}catch(exception){" & _
					vbCrLf & "		alert('" & strTxtErrorInsertingObject & "');" & _
					vbCrLf & "		editor.contentWindow.focus();" & _
					vbCrLf & "	}")
	End If
	
	Response.Write(vbCrLf & "}")
	
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

//Function to preview URL
function showPreview(attached){

	//Get the file extension to check
	var extension = attached;
	extension.toLowerCase();
	extension = extension.slice(extension.lastIndexOf('.'), extension.length);
	
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
			document.getElementById("prevWindow").contentWindow.location.href = attached;
		}catch(exception){
		}
	
	}else{
		document.getElementById("prevWindow").contentWindow.location.href="RTE_popup_link_preview.asp?b=0<% = strQsSID2 %>";
	
	}
}

//Function to check upload file is selected
function checkFile(){
	if (document.getElementById('file').value==''){
	
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
      <td colspan="2" width="63%"><h1><% = strTxAttachFileProperties %></h1></td>
    </tr>
    <tr>
      <td colspan="2" class="RTEtableRow"><table width="100%" border="0" align="center" cellpadding="1" cellspacing="0">
        <tr>
          <td width="42%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="2">
            <tr>
              <td width="88%" class="text"><% = strTxtPath %>: <span id="path"><% = strAttachFilePath %>" </span></td>
            </tr>
            <tr>
              <td class="text"><% = strTxtFileName %>:<iframe src="RTE_file_browser.asp<% = strQsSID1 %>" id="fileWindow" width="98%" height="180px" style="border: #A5ACB2 1px solid;background-color: #FFFFFF;"></iframe></td>
            </tr>
            <form method="post" action="RTE_popup_file_atch.asp?PB=Y<% = strAspUploadPID %><% = strQsSID2 %>" name="frmUpload" enctype="multipart/form-data" onsubmit="return checkFile();" >
             <tr>
              <td class="text"><% = strTxtFileUpload %></td>
            </tr>
            <tr>
              <td class="smText"><% Response.Write(strTxtFilesMustBeOfTheType & ", " & Replace(strUploadFileTypes, ";", ", ", 1, -1, 1) & ", " & strTxtAndHaveMaximumFileSizeOf & " " & strMaxFileUpload) %></td>
            </tr>
            <tr>
              <td><input id="file" name="file" type="file" size="35" /></td>
            </tr>
            <tr>
              <td>
              	<input name="upload" type="submit" id="upload" value="Upload">
              </td>
            </tr>
           </form>
          </table></td>
          <td width="58%" valign="top">
          <form method="post" action="RTE_popup_file_atch.asp<% = strQsSID1 %>" name="frmImageInsrt">
            <table width="100%" border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td width="19%" align="right" class="text"><% = strTxtFileURL %>:</td>
                <td width="81%" colspan="5"><input name="URL" type="text" id="URL" value="" size="40" onchange="document.getElementById('Submit').disabled=false;" onkeypress="document.getElementById('Submit').disabled=false;">
                  <input name="preview" type="button" id="preview" value="<% = strTxtPreview %>" onclick="showPreview(document.getElementById('URL').value)">
                </td>
                <tr>
                <td align="right" class="text"><% = strTxtTitle %>:</td>
                <td><input name="Title" type="text" id="Title" size="40" maxlength="40"></td>
              </tr>
              <tr>
                <td align="right" class="text"><% = strTxtWindow %>:</td>
                <td><select name="windowSel" id="windowSel" onchange="document.getElementById('Window').value=this.options[this.selectedIndex].value">
                  <option value="" selected>Default</option>
                  <option value="_blank">New Window</option>
                  <option value="_self">Same Window</option>
                  <option value="_parent">Parent Window</option>
                  <option value="_top">Top Window</option>
                </select>
                  <input name="Window" type="text" id="Window" size="12" maxlength="15"></td>
              </tr>
              </tr>
                <td align="right" valign="top" class="text"><% = strTxtPreview %>:</td>
                <td colspan="5"><iframe src="RTE_popup_link_preview.asp<% = strQsSID1 %>" id="prevWindow" width="98%" height="220px" style="border: #A5ACB2 1px solid;background-color: #FFFFFF;"></iframe></td>
                </tr>
          </table></td>
        </tr>
      </table></td>
    </tr>
    <tr>
    <td class="RTEtableBottomRow" valign="top">&nbsp;<%

'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnAbout Then
	Response.Write("<span class=""text"" style=""font-size:10px""><a href=""http://www.richtexteditor.org"" target=""_blank"" style=""font-size:10px"">Web Wiz Rich Text Editor</a> version " & strRTEversion & "</span>")
End If 
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******      
      
      %></td>
      <td align="right" class="RTEtableBottomRow"><input type="hidden" name="postBack" value="true">
          <input type="submit" name="Submit" id="Submit" value="     <% = strTxtOK %>     ">&nbsp;<input type="button" name="cancel" value=" <% = strTxtCancel %> " onclick="window.close()">
        </form>
        <br /><br />
      </td>
    </tr>

</table>
</body>
</html><%

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