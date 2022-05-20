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


'Dimension variables
Dim strMode		'Holds the mode of the page, set to true if changes are to be made to the database
Dim strFilePath		'Holds the path to the files
Dim saryBadFileTypes(60)'Array for bad file types
Dim blnBadFileType	'Found bad file type
Dim intLoopCounter	'Loop counter
Dim intLoopCounter2	'Loop counter
Dim strBadFileTypeName	'For error message
Dim saryImageFileType	'Array holding the file types




blnBadFileType = false


'Read in the details from the form
strUploadComponent = Request.Form("component")
strImageTypes = Request.Form("imageTypes")
intUploadAllocatedSpace = Request.Form("allocatedSpace")
lngUploadMaxImageSize = LngC(Request.Form("imageSize"))
strUploadFileTypes = Request.Form("fileTypes")
lngUploadMaxFileSize = LngC(Request.Form("fileSize"))
strAvatarTypes = Request.Form("avatarTypes")
intMaxAvatarSize = IntC(Request.Form("avatarSize"))
blnAvatarUploadEnabled = BoolC(Request.Form("avatar"))
blnUploadSecurityCheck = BoolC(Request.Form("scan"))
intMaxImageWidth = IntC(Request.Form("imgWidth"))
intMaxImageHeight = IntC(Request.Form("imgHeight"))
blnUploadDuplicateFiles = BoolC(Request.Form("duplicateFiles"))

	





'If a hacker gains control of the admin account they can use the upload tool to upload files to the server to hack the entire site
'To prevent this certain file types are not allowed
If Request.Form("postBack") Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))

	'List of bad file types
	
	'ISAPI and CGI web page extensions (can be used to hack site)
	saryBadFileTypes(0) = "asax"
	saryBadFileTypes(1) = "ascx"
	saryBadFileTypes(2) = "ashx"
	saryBadFileTypes(3) = "asmx"
	saryBadFileTypes(4) = "aspx"
	saryBadFileTypes(5) = "asp"
	saryBadFileTypes(6) = "asa"
	saryBadFileTypes(7) = "asr"
	saryBadFileTypes(8) = "axd"
	saryBadFileTypes(9) = "cdx"
	saryBadFileTypes(10) = "cer"
	saryBadFileTypes(11) = "cgi"
	saryBadFileTypes(12) = "class"
	saryBadFileTypes(13) = "config"
	saryBadFileTypes(14) = "com" 
	saryBadFileTypes(15) = "cs"
	saryBadFileTypes(16) = "csproj"
	saryBadFileTypes(17) = "cnf"
	saryBadFileTypes(18) = "dll"
	saryBadFileTypes(19) = "edml"
	saryBadFileTypes(20) = "exe"
	saryBadFileTypes(21) = "idc"
	saryBadFileTypes(22) = "inc"
	saryBadFileTypes(23) = "isp"
	saryBadFileTypes(24) = "licx"
	saryBadFileTypes(25) = "php3"
	saryBadFileTypes(26) = "php4"
	saryBadFileTypes(27) = "php5"
	saryBadFileTypes(28) = "php"
	saryBadFileTypes(29) = "phtml"
	saryBadFileTypes(30) = "pl"
	saryBadFileTypes(31) = "rem"
	saryBadFileTypes(32) = "resources"
	saryBadFileTypes(33) = "resx"
	saryBadFileTypes(34) = "shtm"
	saryBadFileTypes(35) = "shtml"
	saryBadFileTypes(36) = "soap"
	saryBadFileTypes(37) = "stm"
	saryBadFileTypes(38) = "vsdisco"
	saryBadFileTypes(39) = "vbe"
	saryBadFileTypes(40) = "vbs"
	saryBadFileTypes(41) = "vbx"
	saryBadFileTypes(42) = "vb"
	saryBadFileTypes(43) = "webinfo"
	saryBadFileTypes(44) = "cfm"
	saryBadFileTypes(45) = "ssi"
	saryBadFileTypes(46) = "swf"
	saryBadFileTypes(47) = "vbs"
	saryBadFileTypes(48) = "tpl"
	saryBadFileTypes(49) = "cfc"
	saryBadFileTypes(50) = "jst"
	saryBadFileTypes(51) = "jsp"
	saryBadFileTypes(52) = "jse"
	saryBadFileTypes(53) = "jsf"
	saryBadFileTypes(54) = "js"
	saryBadFileTypes(55) = "java"
	saryBadFileTypes(56) = "wml"
	saryBadFileTypes(57) = "xslt"
	saryBadFileTypes(58) = "ini"
	saryBadFileTypes(59) = "htaccess"
	saryBadFileTypes(60) = "osp"
	
	
	'Remove spaces and dots in file types
	strUploadFileTypes = Replace(strUploadFileTypes, " ", "", 1, -1, 1)
	strUploadFileTypes = Replace(strUploadFileTypes, ".", "", 1, -1, 1)
	strImageTypes = Replace(strImageTypes, " ", "", 1, -1, 1)
	strImageTypes = Replace(strImageTypes, ".", "", 1, -1, 1)
	strAvatarTypes = Replace(strAvatarTypes, " ", "", 1, -1, 1)
	strAvatarTypes = Replace(strAvatarTypes, ".", "", 1, -1, 1)
	
	
	'Place the file and image types into an array
	saryImageFileType = Split(Trim(strImageTypes) & ";" & Trim(strUploadFileTypes) & ";" & Trim(strAvatarTypes), ";")
	
	'Loop through all the allowed extensions and see if the image has one
	For intLoopCounter = 0 To UBound(saryImageFileType)
	
		'Loop through each of the file types
		For intLoopCounter2 = 0 To UBound(saryBadFileTypes)
	
			'Check to see if the image extension is allowed
			If LCase(saryImageFileType(intLoopCounter)) = LCase(saryBadFileTypes(intLoopCounter2)) Then 
				blnBadFileType = True
				strBadFileTypeName = strBadFileTypeName & saryBadFileTypes(intLoopCounter2)& ", "
			End If
		Next
	Next
End If




'If the user is changing the upload setup then update the database
If Request.Form("postBack") AND blnBadFileType = false AND blnDemoMode = False Then

	
	Call addConfigurationItem("Upload_component", strUploadComponent)
	Call addConfigurationItem("Upload_img_types", strImageTypes)
	Call addConfigurationItem("Upload_img_size", lngUploadMaxImageSize)
	Call addConfigurationItem("Upload_files_type", strUploadFileTypes)
	Call addConfigurationItem("Upload_files_size", lngUploadMaxFileSize)
	Call addConfigurationItem("Upload_avatar_types", strAvatarTypes)
	Call addConfigurationItem("Upload_avatar_size", intMaxAvatarSize)
	Call addConfigurationItem("Upload_avatar", blnAvatarUploadEnabled)
	Call addConfigurationItem("Upload_allocation", intUploadAllocatedSpace)
	Call addConfigurationItem("Upload_file_scan", blnUploadSecurityCheck)
	Call addConfigurationItem("Upload_img_width", intMaxImageWidth)
	Call addConfigurationItem("Upload_img_height", intMaxImageHeight)
	Call addConfigurationItem("Upload_Duplicate_Files", blnUploadDuplicateFiles)
	

		
	
	'Empty the application level variable so that the changes made are seen in the main forum
	Application.Lock
	
	Application(strAppPrefix & "strUploadComponent") = strUploadComponent
	Application(strAppPrefix & "strImageTypes") = strImageTypes
	Application(strAppPrefix & "lngUploadMaxImageSize") = CLng(lngUploadMaxImageSize)
	Application(strAppPrefix & "strUploadFileTypes") = strUploadFileTypes
	Application(strAppPrefix & "lngUploadMaxFileSize") = CLng(lngUploadMaxFileSize)
	Application(strAppPrefix & "strAvatarTypes") = strAvatarTypes
	Application(strAppPrefix & "intMaxAvatarSize") = CInt(intMaxAvatarSize)
	Application(strAppPrefix & "lngUploadMaxImageSize") = CLng(lngUploadMaxImageSize)
	Application(strAppPrefix & "intUploadAllocatedSpace") = CLng(intUploadAllocatedSpace)
	Application(strAppPrefix & "blnUploadSecurityCheck") = CBool(blnUploadSecurityCheck)
	Application(strAppPrefix & "intMaxImageWidth") = CInt(intMaxImageWidth)
	Application(strAppPrefix & "intMaxImageHeight") = CInt(intMaxImageHeight)
	Application(strAppPrefix & "blnUploadDuplicateFiles") = CBool(blnUploadDuplicateFiles)
	
	Application(strAppPrefix & "blnConfigurationSet") = false
	Application.UnLock	
	
End If



'Initialise the SQL variable with an SQL statement to get the configuration details from the database
strSQL = "SELECT " & strDbTable & "SetupOptions.Option_Item, " & strDbTable & "SetupOptions.Option_Value " & _
"FROM " & strDbTable & "SetupOptions" &  strDBNoLock & " " & _
"ORDER BY " & strDbTable & "SetupOptions.Option_Item ASC;"

'Query the database
rsCommon.Open strSQL, adoCon

'Read in the deatils from the database
If NOT rsCommon.EOF Then
	
	'Place into an array for performance
	saryConfiguration = rsCommon.GetRows()

	'Read in the e-mail setup from the database
	strUploadComponent = getConfigurationItem("Upload_component", "string")
	strImageTypes = getConfigurationItem("Upload_img_types", "string")
	lngUploadMaxImageSize	= CLng(getConfigurationItem("Upload_img_size", "numeric"))
	strUploadFileTypes = getConfigurationItem("Upload_files_type", "string")
	lngUploadMaxFileSize = CLng(getConfigurationItem("Upload_files_size", "numeric"))
	strAvatarTypes = getConfigurationItem("Upload_avatar_types", "string")
	intMaxAvatarSize = CInt(getConfigurationItem("Upload_avatar_size", "numeric"))
	blnAvatarUploadEnabled = CBool(getConfigurationItem("Upload_avatar", "bool"))
	If CLng(getConfigurationItem("Upload_allocation", "numeric")) = 0 Then intUploadAllocatedSpace = 1 Else intUploadAllocatedSpace = CLng(getConfigurationItem("Upload_allocation", "numeric"))
	blnUploadSecurityCheck = CBool(getConfigurationItem("Upload_file_scan", "bool"))
	intMaxImageWidth = CInt(getConfigurationItem("Upload_img_width", "numeric"))
	intMaxImageHeight = CInt(getConfigurationItem("Upload_img_height", "numeric"))
	blnUploadDuplicateFiles = CBool(getConfigurationItem("Upload_Duplicate_Files", "bool"))
	

End If


'Close db
rsCommon.Close




'Initalise the strSQL variable with an SQL statement to query the database
'WHERE cluse added to get round bug in myODBC which won't run an ADO update unless you have a WHERE cluase
strSQL = "SELECT " & strDbTable & "Group.* " & _
"FROM " & strDbTable & "Group " & _
"WHERE " & strDbTable & "Group.Group_ID > 0 " & _
"ORDER BY " & strDbTable & "Group.Group_ID ASC;"
	
'Set the cursor type property of the record set to Forward Only
rsCommon.CursorType = 0

'Set the Lock Type for the records so that the record set is only locked when it is updated
rsCommon.LockType = 3

'Query the database
rsCommon.Open strSQL, adoCon




'Update the db with file and image upload for groups
If Request.Form("postBack") AND blnBadFileType = false Then
	
	'Loop through cats
	Do While NOT rsCommon.EOF
	
		'Update the recordset
		rsCommon.Fields("Image_uploads") = BoolC(Request.Form("imageGroup" & rsCommon("Group_ID")))
		rsCommon.Fields("File_uploads") = BoolC(Request.Form("fileGroup" & rsCommon("Group_ID")))

		'Update the database
		rsCommon.Update
   
		'Move to next record in rs
		rsCommon.MoveNext
	Loop
	
	
	'Re-run the query to read in the updated recordset from the database
	'.Requery
End If



%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Upload Settings</title>
<meta name="generator" content="Web Wiz Forums" />

<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write("<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Forums(TM) ver. " & strVersion & "" & _
vbCrLf & "Info: http://www.webwizforums.com" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->" & vbCrLf)
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>
<script  language="JavaScript" type="text/javascript">

//Function to check form is filled in correctly before submitting
function CheckForm () {


	//Check for a image types name
	if (document.frmUpload.imageTypes.value==""){
		alert("Please enter Image file types to upload");
		document.frmUpload.imageTypes.focus();
		return false;
	}

	//Check for a file types name
	if (document.frmUpload.fileTypes.value==""){
		alert("Please enter File types to upload");
		document.frmUpload.fileTypes.focus();
		return false;
	}
	
	//Check for a avatar types name
	if (document.frmUpload.avatarTypes.value==""){
		alert("Please enter Avatar types to upload");
		document.frmUpload.avatarTypes.focus();
		return false;
	}

	return true
}<%

'If error display message
If blnBadFileType Then
	
	Response.Write(vbCrLf & "alert('For security the following unsafe file type\(s\) are not permited.\n\n" & strBadFileTypeName & "')")	
End If

%>	
</script>
<style type="text/css">
<!--
.style1 {
	color: #FF0000;
	font-weight: bold;
	font-size: 16px;
}
-->
</style>
<!-- #include file="includes/admin_header_inc.asp" -->
<div align="center">
 <h1>Upload Settings</h1>
 <br />
  <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
  <br />
  <table border="0" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
      <td align="center" class="tableLedger">Important - Please Read First!</td>
    </tr>
    <tr class="tableRow" align="left">
      <td>To be able to use file and image upload in your forums, you must have an upload component installed on the web server, if you are unsure about this check or use the   <a href="admin_server_test.asp">Server Compatibility Test</a> page to see which upload components are installed on the server.<br />
          <br />
          If you run the web server yourself then you could download and install one of the following supported components.<br />
          <br />
          You will also need to make sure that the upload folder and it's subfolders have read, write and modify permissions for the Internet User Account (IUSR_&lt;MachineName&gt;) and is inside the root of your forum.
          <ul>
          <li class="text"><span>Persits AspUpload 3.x or above (use this components for uploads above 2MB (2048KB), also includes progress bar)<br />
Component available form <a href="http://www.aspupload.com" target="_blank">www.aspupload.com</a><br />
         Persits AspUpload</span> 2.x<br />
            Component available form <a href="http://www.aspupload.com" target="_blank">www.aspupload.com</a></li>
         <li class="text"><span>Dundas Upload</span> 2.0<br />
            Free component available from <a href="http://aspalliance.com/dundas/default.aspx" target="_blank">aspalliance.com/dundas</a></li>
          <li class="text"><span>SoftArtisans FileUp</span> 3.2 or above (<span>SA FileUp</span>)<br />
            Component available form <a href="http://www.softartisans.com" target="_blank">www.softartisans.com</a></li>
          <li class="text"><span>aspSmartUpload</span><br />
            Free component (no longer has a website)</li>
          <li class="text"><span>AspSimpleUpload</span><br />
            Free component (no longer has a website)</li>
        </ul>
       <p class="text"><span>Please note</span>:<br />
       	 The ASP<span> File System Object</span> (FSO) is also required when using upload features, use the   <a href="admin_server_test.asp">Server Compatibility Test</a> page to check it is not disabled.<br />
       	 <a href="http://www.aspjpeg.com" target="_blank">Persits AspJPEG</a> needs to be installed on the server in order to re-size images when they are uploaded.
           <br /><br />
                 <strong>Maximum Upload File Sizes</strong><br />
               Only Persists AspUpload 3 or above allows uploads above 2MB (2048KB) in size. The server also needs to be modified to accept larger HTTP Requests to allow large file uploads, by default this is 4MB on Windows Server 2000 and 2003 and 30MB on Windows Server 2008.<br />
          <br />
          <span class="style1">Security Warning</span> <strong>- Best Practice </strong><br />
         Allowing users to upload their own files and images requires that write and modify permissions are enabled on the upload directory for the Internet User Account (IUSR). The best practice for this is to ONLY allow write and modify permissions on the upload directory and 'read only' permissions for the rest of your web site. In the event that your site comes under attack from a hacker who manages to gain control  through the IUSR account, this measure prevents the hacker from destroying or defacing the rest of your web site. </td>
    </tr>
  </table>
</div>
<br />
<form action="admin_upload_configure.asp<% = strQsSID1 %>" method="post" name="frmUpload" id="frmUpload" onsubmit="return CheckForm();">
  <table width="100%" height="182" border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr align="left">
      <td height="30" colspan="2" class="tableLedger">General Upload Setup </td>
    </tr>
    <tr>
      <td align="left" class="tableRow">Upload Component to use:<br />
        <span class="smText">You must have the component you select installed on the web server.<br />You can use the <a href="admin_server_test.asp<% = strQsSID1 %>" class="smLink">Server Compatibility Test Tool</a> to see which components you have installed on the server.</span></td>
      <td valign="top" class="tableRow"><select name="component"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
      	  <option value="AspUpload"<% If strUploadComponent = "AspUpload" Then Response.Write(" selected") %>>Persits AspUpload 3.0 or above</option>
      	  <option value="AspUpload2"<% If strUploadComponent = "AspUpload2" Then Response.Write(" selected") %>>Persits AspUpload 2.0</option>
          <option value="Dundas"<% If strUploadComponent = "Dundas" Then Response.Write(" selected") %>>Dundas Upload</option>
          <option value="fileUp"<% If strUploadComponent = "fileUp" Then Response.Write(" selected") %>>SA FileUp</option>
          <option value="aspSmart"<% If strUploadComponent = "aspSmart" Then Response.Write(" selected") %>>aspSmartUpload</option>
          <option value="AspSimple"<% If strUploadComponent = "AspSimple" Then Response.Write(" selected") %>>AspSimpleUpload</option>
          <option value="NoComponent"<% If strUploadComponent = "NoComponent" Then Response.Write(" selected") %>>Without Component</option>
      </select></td>
    </tr>
    <tr>
      <td width="59%" align="left" class="tableRow">Allocated Upload Space:<br />
      <span class="smText">This is the amount of space allocated to each of your members on the server for uploading files and images to.</span></td>
      <td width="41%" valign="top" class="tableRow"><select name="allocatedSpace" id="allocatedSpace"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option<% If intUploadAllocatedSpace = 1 Then Response.Write(" selected") %> value="1">1 MB</option>
       <option<% If intUploadAllocatedSpace = 2 Then Response.Write(" selected") %> value="2">2 MB</option>
       <option<% If intUploadAllocatedSpace = 3 Then Response.Write(" selected") %> value="3">3 MB</option>
       <option<% If intUploadAllocatedSpace = 4 Then Response.Write(" selected") %> value="4">4 MB</option>
       <option<% If intUploadAllocatedSpace = 5 Then Response.Write(" selected") %> value="5">5 MB</option>
       <option<% If intUploadAllocatedSpace = 10 Then Response.Write(" selected") %> value="10">10 MB</option>
       <option<% If intUploadAllocatedSpace = 20 Then Response.Write(" selected") %> value="20">20 MB</option>
       <option<% If intUploadAllocatedSpace = 25 Then Response.Write(" selected") %> value="25">25 MB</option>
       <option<% If intUploadAllocatedSpace = 30 Then Response.Write(" selected") %> value="30">30 MB</option>
       <option<% If intUploadAllocatedSpace = 40 Then Response.Write(" selected") %> value="40">40 MB</option>
       <option<% If intUploadAllocatedSpace = 50 Then Response.Write(" selected") %> value="50">50 MB</option>
       <option<% If intUploadAllocatedSpace = 60 Then Response.Write(" selected") %> value="60">60 MB</option>
       <option<% If intUploadAllocatedSpace = 70 Then Response.Write(" selected") %> value="70">70 MB</option>
       <option<% If intUploadAllocatedSpace = 80 Then Response.Write(" selected") %> value="80">80 MB</option>
       <option<% If intUploadAllocatedSpace = 90 Then Response.Write(" selected") %> value="90">90 MB</option>
       <option<% If intUploadAllocatedSpace = 100 Then Response.Write(" selected") %> value="100">100 MB</option>
       <option<% If intUploadAllocatedSpace = 125 Then Response.Write(" selected") %> value="125">125 MB</option>
       <option<% If intUploadAllocatedSpace = 150 Then Response.Write(" selected") %> value="150">150 MB</option>
       <option<% If intUploadAllocatedSpace = 175 Then Response.Write(" selected") %> value="175">175 MB</option>
       <option<% If intUploadAllocatedSpace = 200 Then Response.Write(" selected") %> value="200">200 MB</option>
       <option<% If intUploadAllocatedSpace = 250 Then Response.Write(" selected") %> value="250">250 MB</option>
       <option<% If intUploadAllocatedSpace = 500 Then Response.Write(" selected") %> value="500">500 MB</option>
       <option<% If intUploadAllocatedSpace = 1024 Then Response.Write(" selected") %> value="1024">1 GB</option>
       <option<% If intUploadAllocatedSpace = 2048 Then Response.Write(" selected") %> value="2048">2 GB</option>
       <option<% If intUploadAllocatedSpace = 5120 Then Response.Write(" selected") %> value="5120">5 GB</option>
       <option<% If intUploadAllocatedSpace = 10240 Then Response.Write(" selected") %> value="10240">10 GB</option>
       <option<% If intUploadAllocatedSpace = 51200 Then Response.Write(" selected") %> value="51200">50 GB</option>
       <option<% If intUploadAllocatedSpace = 102400 Then Response.Write(" selected") %> value="102400">100 GB</option>
      </select>
      </td>
    </tr>
    
    
   
     <tr>
      <td  height="13" align="left" class="tableRow">Allow Duplicate File Names<br />
      	<span class="smText">When enabled if a file or image of the same name is uploaded it will be renamed rather than rejected. As Apple IOS (IPhone/IPad) name all images 'upload.jpeg' it is recommend to enable this feature.</span></td>
      <td height="13" valign="top" class="tableRow">Yes
        <input type="radio" name="duplicateFiles" value="True" <% If blnUploadDuplicateFiles  Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
        &nbsp;&nbsp;No
        <input type="radio" name="duplicateFiles" value="False" <% If blnUploadDuplicateFiles = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    
    
    <tr>
      <td  height="13" align="left" class="tableRow">Upload Image/File Scanning<br />
      	<span class="smText">Parse uploaded files and images for malicious code that can be used in an XSS Attack against your forum. Images often contain meta data which can lead to some false positives.</span></td>
      <td height="13" valign="top" class="tableRow">Yes
        <input type="radio" name="scan" value="True" <% If blnUploadSecurityCheck  Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
        &nbsp;&nbsp;No
        <input type="radio" name="scan" value="False" <% If blnUploadSecurityCheck = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
      <td colspan="2" align="left" class="tableLedger">Image Upload</td>
    </tr>
    <tr>
      <td width="59%" align="left" class="tableRow">Image Types*<br />
        <span class="smText">Place the types of images that can be uploaded in posts. Separate the different image 
        types with a semi-colon.<br />
      eg. jpg;jpeg;gif;png</span></td>
      <td width="41%" valign="top" class="tableRow"><input name="imageTypes" type="text" id="imageTypes" value="<% = strImageTypes %>" size="50" maxlength="125"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
     <td  height="2" align="left" class="tableRow">Maximum Image File Size<br />
       <span class="smText">This is the maximum file size of images members can upload.</span>
     </td>
     <td height="2" valign="top" class="tableRow"><select name="imageSize" id="imageSize"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option<% If lngUploadMaxImageSize = 10 Then Response.Write(" selected") %> value="10">10 KB</option>
       <option<% If lngUploadMaxImageSize = 20 Then Response.Write(" selected") %> value="20">20 KB</option>
       <option<% If lngUploadMaxImageSize = 30 Then Response.Write(" selected") %> value="30">30 KB</option>
       <option<% If lngUploadMaxImageSize = 40 Then Response.Write(" selected") %> value="40">40 KB</option>
       <option<% If lngUploadMaxImageSize = 50 Then Response.Write(" selected") %> value="50">50 KB</option>
       <option<% If lngUploadMaxImageSize = 60 Then Response.Write(" selected") %> value="60">60 KB</option>
       <option<% If lngUploadMaxImageSize = 80 Then Response.Write(" selected") %> value="80">80 KB</option>
       <option<% If lngUploadMaxImageSize = 100 Then Response.Write(" selected") %> value="100">100 KB</option>
       <option<% If lngUploadMaxImageSize = 150 Then Response.Write(" selected") %> value="150">150 KB</option>
       <option<% If lngUploadMaxImageSize = 200 Then Response.Write(" selected") %> value="200">200 KB</option>
       <option<% If lngUploadMaxImageSize = 300 Then Response.Write(" selected") %> value="300">300 KB</option>
       <option<% If lngUploadMaxImageSize = 400 Then Response.Write(" selected") %> value="400">400 KB</option>
       <option<% If lngUploadMaxImageSize = 500 Then Response.Write(" selected") %> value="500">500 KB</option>
       <option<% If lngUploadMaxImageSize = 1024 Then Response.Write(" selected") %> value="1024">1 MB</option>
       <option<% If lngUploadMaxImageSize = 2048 Then Response.Write(" selected") %> value="2048">2 MB</option>
       <option<% If lngUploadMaxImageSize = 3072 Then Response.Write(" selected") %> value="3072">3 MB</option>
       <option<% If lngUploadMaxImageSize = 4096 Then Response.Write(" selected") %> value="4096">4 MB</option>
       <option<% If lngUploadMaxImageSize = 5120 Then Response.Write(" selected") %> value="5120">5 MB</option>
       <option<% If lngUploadMaxImageSize = 7168 Then Response.Write(" selected") %> value="7168">7 MB</option>
       <option<% If lngUploadMaxImageSize = 10240 Then Response.Write(" selected") %> value="10240">10 MB</option>
       <option<% If lngUploadMaxImageSize = 15360 Then Response.Write(" selected") %> value="15360">15 MB</option>
       <option<% If lngUploadMaxImageSize = 20480 Then Response.Write(" selected") %> value="20480">20 MB</option>
       <option<% If lngUploadMaxImageSize = 30720 Then Response.Write(" selected") %> value="30720">30 MB</option>
       <option<% If lngUploadMaxImageSize = 40960 Then Response.Write(" selected") %> value="40960">40 MB</option>
       <option<% If lngUploadMaxImageSize = 51200 Then Response.Write(" selected") %> value="51200">50 MB</option>
      </select>
      </td>
    </tr>
    
     <tr>
     <td  height="2" align="left" class="tableRow">Maximum Image Width<br />
       <span class="smText">This is the maximum width an uploaded image can be, if the image is larger it will be re-sized to this size.<br />
       	You must have Persits AspJPEG inststalled on the server to use this. You can use the <a href="admin_server_test.asp<% = strQsSID1 %>" class="smLink">Server Compatibility Test Tool</a> to see which components you have installed on the server.</span>
     </td>
     <td height="2" valign="top" class="tableRow"><select name="imgWidth" id="imgWidth"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option<% If intMaxImageWidth = 0 Then Response.Write(" selected") %> value="0">Off</option>
       <option<% If intMaxImageWidth = 50 Then Response.Write(" selected") %> value="50">50 pixels</option>
       <option<% If intMaxImageWidth = 75 Then Response.Write(" selected") %> value="75">75 pixels</option>
       <option<% If intMaxImageWidth = 100 Then Response.Write(" selected") %> value="100">100 pixels</option>
       <option<% If intMaxImageWidth = 150 Then Response.Write(" selected") %> value="150">150 pixels</option>
       <option<% If intMaxImageWidth = 200 Then Response.Write(" selected") %> value="200">200 pixels</option>
       <option<% If intMaxImageWidth = 250 Then Response.Write(" selected") %> value="250">250 pixels</option>
       <option<% If intMaxImageWidth = 300 Then Response.Write(" selected") %> value="300">300 pixels</option>
       <option<% If intMaxImageWidth = 350 Then Response.Write(" selected") %> value="350">350 pixels</option>
       <option<% If intMaxImageWidth = 400 Then Response.Write(" selected") %> value="400">400 pixels</option>
       <option<% If intMaxImageWidth = 450 Then Response.Write(" selected") %> value="450">450 pixels</option>
       <option<% If intMaxImageWidth = 500 Then Response.Write(" selected") %> value="500">500 pixels</option>
       <option<% If intMaxImageWidth = 600 Then Response.Write(" selected") %> value="600">600 pixels</option>
       <option<% If intMaxImageWidth = 700 Then Response.Write(" selected") %> value="700">700 pixels</option>
       <option<% If intMaxImageWidth = 800 Then Response.Write(" selected") %> value="800">800 pixels</option>
       <option<% If intMaxImageWidth = 900 Then Response.Write(" selected") %> value="900">900 pixels</option>
       <option<% If intMaxImageWidth = 1000 Then Response.Write(" selected") %> value="1000">1000 pixels</option>
       <option<% If intMaxImageWidth = 1100 Then Response.Write(" selected") %> value="1100">1100 pixels</option>
       <option<% If intMaxImageWidth = 1200 Then Response.Write(" selected") %> value="1200">1200 pixels</option>
       <option<% If intMaxImageWidth = 1300 Then Response.Write(" selected") %> value="1300">1300 pixels</option>
       <option<% If intMaxImageWidth = 1400 Then Response.Write(" selected") %> value="1400">1400 pixels</option>
       <option<% If intMaxImageWidth = 1500 Then Response.Write(" selected") %> value="1500">1500 pixels</option>
       <option<% If intMaxImageWidth = 1600 Then Response.Write(" selected") %> value="1500">1500 pixels</option>
       <option<% If intMaxImageWidth = 2000 Then Response.Write(" selected") %> value="2000">2000 pixels</option>
       <option<% If intMaxImageWidth = 2500 Then Response.Write(" selected") %> value="2500">2500 pixels</option>
      </select>
      </td>
    </tr>
     <tr>
     <td  height="2" align="left" class="tableRow">Maximum Image Height<br />
        <span class="smText">This is the maximum height an uploaded image can be, if the image is larger it will be re-sized to this size.<br />
       	You must have Persits AspJPEG inststalled on the server to use this. You can use the <a href="admin_server_test.asp<% = strQsSID1 %>" class="smLink">Server Compatibility Test Tool</a> to see which components you have installed on the server.</span>
     </td>
     <td height="2" valign="top" class="tableRow"><select name="imgHeight" id="imgHeight"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option<% If intMaxImageHeight = 0 Then Response.Write(" selected") %> value="0">Off</option>
       <option<% If intMaxImageHeight = 50 Then Response.Write(" selected") %> value="50">50 pixels</option>
       <option<% If intMaxImageHeight = 75 Then Response.Write(" selected") %> value="75">75 pixels</option>
       <option<% If intMaxImageHeight = 100 Then Response.Write(" selected") %> value="100">100 pixels</option>
       <option<% If intMaxImageHeight = 150 Then Response.Write(" selected") %> value="150">150 pixels</option>
       <option<% If intMaxImageHeight = 200 Then Response.Write(" selected") %> value="200">200 pixels</option>
       <option<% If intMaxImageHeight = 250 Then Response.Write(" selected") %> value="250">250 pixels</option>
       <option<% If intMaxImageHeight = 300 Then Response.Write(" selected") %> value="300">300 pixels</option>
       <option<% If intMaxImageHeight = 350 Then Response.Write(" selected") %> value="350">350 pixels</option>
       <option<% If intMaxImageHeight = 400 Then Response.Write(" selected") %> value="400">400 pixels</option>
       <option<% If intMaxImageHeight = 450 Then Response.Write(" selected") %> value="450">450 pixels</option>
       <option<% If intMaxImageHeight = 500 Then Response.Write(" selected") %> value="500">500 pixels</option>
       <option<% If intMaxImageHeight = 600 Then Response.Write(" selected") %> value="600">600 pixels</option>
       <option<% If intMaxImageHeight = 700 Then Response.Write(" selected") %> value="700">700 pixels</option>
       <option<% If intMaxImageHeight = 800 Then Response.Write(" selected") %> value="800">800 pixels</option>
       <option<% If intMaxImageHeight = 900 Then Response.Write(" selected") %> value="900">900 pixels</option>
       <option<% If intMaxImageHeight = 1000 Then Response.Write(" selected") %> value="1000">1000 pixels</option>
       <option<% If intMaxImageHeight = 1100 Then Response.Write(" selected") %> value="1100">1100 pixels</option>
       <option<% If intMaxImageHeight = 1200 Then Response.Write(" selected") %> value="1200">1200 pixels</option>
       <option<% If intMaxImageHeight = 1300 Then Response.Write(" selected") %> value="1300">1300 pixels</option>
       <option<% If intMaxImageHeight = 1400 Then Response.Write(" selected") %> value="1400">1400 pixels</option>
       <option<% If intMaxImageHeight = 1500 Then Response.Write(" selected") %> value="1500">1500 pixels</option>
       <option<% If intMaxImageHeight = 1600 Then Response.Write(" selected") %> value="1500">1500 pixels</option>
       <option<% If intMaxImageHeight = 2000 Then Response.Write(" selected") %> value="2000">2000 pixels</option>
       <option<% If intMaxImageHeight = 2500 Then Response.Write(" selected") %> value="2500">2500 pixels</option>
      </select>
      </td>
    </tr>
    
    
    <tr>
     <td  height="2" colspan="2" align="left" class="tableRow">Select Which Groups are Permitted to Upload Images
      <table width="100%"  border="0" cellspacing="1" cellpadding="1">
      <tr class="tableRow"> 
       <td width="1%" align="right"><input type="checkbox" name="chkAllimageGroup" id="chkAllimageGroup" onclick="checkAll('imageGroup');" /></td>
       <td width="99%"><strong>Check All</strong></td>
      </tr><%
 
'Query the database
rsCommon.MoveFirst        
	
'Loop through cats
Do While NOT rsCommon.EOF
	
	'If not guest group display group to be selected for uploading (you would be stupid to allow a security risk like uploading by guests!!)
	If rsCommon("Group_ID") <> 2 Then
		Response.Write(vbCrLf & "   <tr class=""tableRow""> " & _
		vbCrLf & "    <td width=""1%"" align=""right""><input type=""checkbox"" name=""imageGroup" & rsCommon("Group_ID") & """ id=""imageGroup" & rsCommon("Group_ID") & """ value=""true""")
		If  CBool(rsCommon("Image_uploads")) Then Response.Write(" checked")
		If blnDemoMode Then Response.Write(" disabled=""disabled""")
		Response.Write(" /></td>" & _
		vbCrLf & "    <td width=""99%"">" & rsCommon("Name") & "</td>" & _
		vbCrLf & "   </tr>")
	End If
   
	'Move to next record in rs
	rsCommon.MoveNext
Loop

 %>
       </table>     </td>
     </tr>
    <tr>
      <td  height="7" colspan="2" align="left" class="tableLedger">File Upload</td>
    </tr>
    <tr>
      <td width="59%"  height="13" align="left" class="tableRow">File Types*<br />
        <span class="smText">Place the types of files that can be upload in posts. Separate the different file types with a semi-colon.<br />
      eg. zip;rar</span></td>
      <td width="41%" height="13" valign="top" class="tableRow"><input name="fileTypes" type="text" id="fileTypes" value="<% = strUploadFileTypes %>" size="50" maxlength="125"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
     <td  height="13" align="left" class="tableRow">Maximum File Size<br />
       <span class="smText">This is the maximum file size that members can upload.</span></td>
     <td height="13" valign="top" class="tableRow"><select name="fileSize" id="fileSize"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option<% If lngUploadMaxFileSize = 10 Then Response.Write(" selected") %> value="10">10 KB</option>
       <option<% If lngUploadMaxFileSize = 20 Then Response.Write(" selected") %> value="20">20 KB</option>
       <option<% If lngUploadMaxFileSize = 30 Then Response.Write(" selected") %> value="30">30 KB</option>
       <option<% If lngUploadMaxFileSize = 40 Then Response.Write(" selected") %> value="40">40 KB</option>
       <option<% If lngUploadMaxFileSize = 50 Then Response.Write(" selected") %> value="50">50 KB</option>
       <option<% If lngUploadMaxFileSize = 60 Then Response.Write(" selected") %> value="60">60 KB</option>
       <option<% If lngUploadMaxFileSize = 80 Then Response.Write(" selected") %> value="80">80 KB</option>
       <option<% If lngUploadMaxFileSize = 100 Then Response.Write(" selected") %> value="100">100 KB</option>
       <option<% If lngUploadMaxFileSize = 150 Then Response.Write(" selected") %> value="150">150 KB</option>
       <option<% If lngUploadMaxFileSize = 200 Then Response.Write(" selected") %> value="200">200 KB</option>
       <option<% If lngUploadMaxFileSize = 300 Then Response.Write(" selected") %> value="300">300 KB</option>
       <option<% If lngUploadMaxFileSize = 400 Then Response.Write(" selected") %> value="400">400 KB</option>
       <option<% If lngUploadMaxFileSize = 500 Then Response.Write(" selected") %> value="500">500 KB</option>
       <option<% If lngUploadMaxFileSize = 1024 Then Response.Write(" selected") %> value="1024">1 MB</option>
       <option<% If lngUploadMaxFileSize = 1536 Then Response.Write(" selected") %> value="1536">1.5 MB</option>
       <option<% If lngUploadMaxFileSize = 2048 Then Response.Write(" selected") %> value="2048">2 MB</option>
       <option<% If lngUploadMaxFileSize = 3072 Then Response.Write(" selected") %> value="3072">3 MB</option>
       <option<% If lngUploadMaxFileSize = 4096 Then Response.Write(" selected") %> value="4096">4 MB</option>
       <option<% If lngUploadMaxFileSize = 5120 Then Response.Write(" selected") %> value="5120">5 MB</option>
       <option<% If lngUploadMaxFileSize = 7168 Then Response.Write(" selected") %> value="7168">7 MB</option>
       <option<% If lngUploadMaxFileSize = 10240 Then Response.Write(" selected") %> value="10240">10 MB</option>
       <option<% If lngUploadMaxFileSize = 15360 Then Response.Write(" selected") %> value="15360">15 MB</option>
       <option<% If lngUploadMaxFileSize = 20480 Then Response.Write(" selected") %> value="20480">20 MB</option>
       <option<% If lngUploadMaxFileSize = 30720 Then Response.Write(" selected") %> value="30720">30 MB</option>
       <option<% If lngUploadMaxFileSize = 40960 Then Response.Write(" selected") %> value="40960">40 MB</option>
       <option<% If lngUploadMaxFileSize = 51200 Then Response.Write(" selected") %> value="51200">50 MB</option>
       <option<% If lngUploadMaxFileSize = 102400 Then Response.Write(" selected") %> value="102400">100 MB</option>
       <option<% If lngUploadMaxFileSize = 204800 Then Response.Write(" selected") %> value="204800">200 MB</option>
       <option<% If lngUploadMaxFileSize = 307200 Then Response.Write(" selected") %> value="307200">300 MB</option>
       <option<% If lngUploadMaxFileSize = 409600 Then Response.Write(" selected") %> value="409600">400 MB</option>
       <option<% If lngUploadMaxFileSize = 512000 Then Response.Write(" selected") %> value="512000">500 MB</option>
      </select>
      </td>
    </tr>
    <tr>
     <td  height="13" colspan="2" align="left" class="tableRow">Select Which Groups are Permitted to Upload Files
     <table width="100%"  border="0" cellspacing="1" cellpadding="1">
     <tr class="tableRow"> 
       <td width="1%" align="right"><input type="checkbox" name="chkAllfileGroup" id="chkAllfileGroup" onclick="checkAll('fileGroup');" /></td>
       <td width="99%"><strong>Check All</strong></td>
      </tr><%
	
'Query the database
rsCommon.MoveFirst

	
'Loop through cats
Do While NOT rsCommon.EOF
	
	'If not guest group display group to be selected for uploading (you would be stupid to allow a security risk like uploading by guests!!)
	If rsCommon("Group_ID") <> 2 Then
		Response.Write(vbCrLf & "   <tr class=""tableRow""> " & _
		vbCrLf & "    <td width=""1%"" align=""right""><input type=""checkbox"" name=""fileGroup" & rsCommon("Group_ID") & """ id=""fileGroup" & rsCommon("Group_ID") & """ value=""true""")
		If  CBool(rsCommon("File_uploads")) Then Response.Write(" checked")
		If blnDemoMode Then Response.Write(" disabled=""disabled""")
		Response.Write(" /></td>" & _
		vbCrLf & "    <td width=""99%"">" & rsCommon("Name") & "</td>" & _
		vbCrLf & "   </tr>")
	End If
   
	'Move to next record in rs
	rsCommon.MoveNext
Loop


'Reset Server Variables
rsCommon.close
 %>
       </table> 
      </td>
     </tr>
    <tr>
      <td  height="13" colspan="2" align="left" class="tableLedger">Avatar Upload</td>
    </tr>
    <tr class="tableSubLedger">
      <td  height="13" colspan="2" align="left"><span class="smText">Make sure you have also enabled Avatar Images from the <a href="admin_user_settings.asp<% = strQsSID1 %>" class="smLink">User Settings</a> page.<br />
            <strong>For extra security avatars can only be uploaded once a user is registered, by editing their profile.</strong></span></td>
    </tr>
    <tr>
      <td  height="13" align="left" class="tableRow">Enable Avatar Uploading</td>
      <td height="13" valign="top" class="tableRow">Yes
        <input type="radio" name="avatar" value="True" <% If blnAvatarUploadEnabled  Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
        &nbsp;&nbsp;No
        <input type="radio" name="avatar" value="False" <% If blnAvatarUploadEnabled = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
      <td  height="13" align="left" class="tableRow">Avatar Image Types*<br />
        <span class="smText">Place the types of images that can be uploaded in posts. Separate the different image types with a semi-colon.<br />
      eg. jpg;jpeg;gif;png</span></td>
      <td height="13" valign="top" class="tableRow"><input name="avatarTypes" type="text" id="avatarTypes" value="<% = strAvatarTypes %>" size="50" maxlength="125"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
      <td  height="13" align="left" class="tableRow">Maximum Avatar Image File Size<br />
      <span class="smText">This is the maximum file size of images members can upload.</span></td>
      <td height="13" valign="top" class="tableRow">
       <select name="avatarSize" id="avatarSize"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
        <option<% If intMaxAvatarSize = 10 Then Response.Write(" selected") %> value="10">10 KB</option>
        <option<% If intMaxAvatarSize = 20 Then Response.Write(" selected") %> value="20">20 KB</option>
        <option<% If intMaxAvatarSize = 30 Then Response.Write(" selected") %> value="30">30 KB</option>
        <option<% If intMaxAvatarSize = 40 Then Response.Write(" selected") %> value="40">40 KB</option>
        <option<% If intMaxAvatarSize = 50 Then Response.Write(" selected") %> value="50">50 KB</option>
        <option<% If intMaxAvatarSize = 60 Then Response.Write(" selected") %> value="60">60 KB</option>
        <option<% If intMaxAvatarSize = 80 Then Response.Write(" selected") %> value="80">80 KB</option>
        <option<% If intMaxAvatarSize = 100 Then Response.Write(" selected") %> value="100">100 KB</option>
        <option<% If intMaxAvatarSize = 150 Then Response.Write(" selected") %> value="150">150 KB</option>
        <option<% If intMaxAvatarSize = 200 Then Response.Write(" selected") %> value="200">200 KB</option>
        <option<% If intMaxAvatarSize = 300 Then Response.Write(" selected") %> value="300">300 KB</option>
        <option<% If intMaxAvatarSize = 400 Then Response.Write(" selected") %> value="400">400 KB</option>
        <option<% If intMaxAvatarSize = 500 Then Response.Write(" selected") %> value="500">500 KB</option>
       </select>
      </td>
    </tr>
    <tr align="center">
      <td height="2" colspan="2" valign="top" class="tableBottomRow" >
          <input type="hidden" name="postBack" value="true" />
          <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
          <input type="submit" name="Submit" value="Update Details" />
          <input type="reset" name="Reset" value="Reset Form" />      </td>
    </tr>
  </table>
</form>
<br />
<!-- #include file="includes/admin_footer_inc.asp" --><%

Call closeDatabase()

%>