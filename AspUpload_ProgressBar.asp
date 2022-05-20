<% @ EnableSessionState = False
Language=VBScript
%>
<% Option Explicit %>
<!-- #include file="language_files/RTE_language_file_inc.asp" -->
<!-- #include file="functions/functions_filters.asp" -->
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





Response.AddHeader "pragma","cache"
Response.AddHeader "cache-control","public"
Response.CacheControl =	"Public"
Response.Expires = -1 



Dim strPID
Dim objUploadProgress
Dim strProgressBar
Dim strProgressBarFormat
Dim intRefeshTime

  
'Read in the PID
strPID = removeAllTags(Request.QueryString("PID"))
intRefeshTime = removeAllTags(Request.QueryString("to"))


'Make sure that we have a PID
If strPID <> "" Then
	
	'Progress Bar format (see http://www.aspupload.com/manual_progress.html for different format options)
	strProgressBarFormat = "%T%P " & strTxtOfFilesUploadedToRemoteServer & "%t%B0%T " & strTxtEstimatedTimeLeft & ": %d %R (%U of %V " & strTxtCopied & ") %l " & strTxtTransferRate & ": %d %S/Sec %t"

	'Create upload progress object
	Set objUploadProgress = Server.CreateObject("Persits.UploadProgress")
	
	'Progress bar
	strProgressBar = objUploadProgress.FormatProgress(strPID, intRefeshTime, "#00007F", strProgressBarFormat)
	
	'Clean up
	Set objUploadProgress = Nothing
End If


'If progress bar is empty then upload complete or timedout
If strProgressBar = "" Then
%>
<html>
<head>
<title>Upload Finished</title>
<script language="JavaScript">
function CloseMe()
{
	window.parent.close();
	return true;
}
</script>
</head>
<body onload="CloseMe()">
</body>
</html><%




'Not finished yet
Else    
%>
<html>
<head>
<meta HTTP-EQUIV="Refresh" CONTENT="1;URL=<%

 	Response.Write(Request.ServerVariables("URL"))
 	Response.Write("?to=" & intRefeshTime & "&PID=" & strPID)
 
  %>" />
<title>Uploading Files...</title>
<style type="text/css">
html, body {
  background: ButtonFace;
  color: ButtonText;
  font: font-family: Verdana, Arial, Helvetica, sans-serif;
  font-size: 12px;
  margin: 0px;
  padding: 0px;
}
td {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 12px;
	color: #000000;
}
td.spread {
	font-size: 6pt; line-height:6pt 
} 
td.brick {
	font-size:6pt; height:15px
}
</style>
</head>
<body>
<% = strProgressBar %>
</body>
</html><% 

End If 

%>