<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="admin_common.asp" -->
<!--#include file="functions/functions_common.asp" -->
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




'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******	

Dim objXMLHTTP, objXmlDoc
Dim intResponseCode
Dim strNewVersionNumber
Dim strReleaseDate
Dim strReleaseAbout
Dim strUpdateServerError
Dim strDataStream
Dim strNewsFeed
Dim strLicenseType
Dim strFID, strFID2


strUpdateServerError = ""
If strInstallID = "" Then strInstallID = "Free Express Edition Install"
strFID = decodeString(strCodeField)
strFID2 = decodeString(strCodeField2)

	

	
'Set objXMLHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
Set objXMLHTTP = Server.CreateObject("Microsoft.XMLHTTP")
On Error Resume Next

'URL to post to and get XML reponse
objXMLHTTP.Open "POST", "http://www.webwizforums.com/update_check.asp", False

'Set headers
objXMLHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objXMLHTTP.setRequestHeader "User-Agent", "WebWizFourms/" & strVersion & "; (ForumURL " & strForumPath  & ")"

'Post data tto server
objXMLHTTP.Send("IP=" & Request.ServerVariables("LOCAL_ADDR") & "&VER=" & strVersion & "&ID=" & strInstallID)
       
'If error then let user know
If Err.Number <> 0 Then strUpdateServerError = "<strong>Fail:</strong><br />Error Connecting to Update Server. <br /><br />If Web Wiz Forums is running on a server behind a Firewall check that TCP Port 80 is open and not using a proxy server."

'If not 200 then bad response
If NOT objXMLHTTP.Status = 200 Then
  	strUpdateServerError = "<strong>Fail:</strong><br />Error Connecting to Update Server. <br /><br />Server Response: " & objXMLHTTP.Status & " - " & objXMLHTTP.statusText & "<br /><br />If Web Wiz Forums is running on a server behind a Firewall check that TCP Port 80 is open is open and not using a proxy server."
	On Error goto 0
	Set objXMLHTTP = Nothing

'Else we have a 200 OK repsonse so parse XML
Else
  	strDataStream = objXMLHTTP.ResponseText
	On Error goto 0
        Set objXMLHTTP = Nothing
       
        'Read in XML
        Set objXmlDoc = CreateObject("Msxml2.FreeThreadedDOMDocument")
	objXmlDoc.Async = False
	objXmlDoc.LoadXML(strDataStream)
	
	'If XML parse fails tell user
	If objXmlDoc.parseError.errorCode <> 0 Then 
		strUpdateServerError = "<strong>Fail:</strong><br />XML Parse Error: " & objXmlDoc.parseError.reason & "<br /><br />If Web Wiz Forums is running on a server behind a Firewall check that TCP Port 80 is open."
      	
      	'Else get data from XML
      	Else
		intResponseCode = CInt(objXmlDoc.childNodes(1).childNodes(0).text)
		
		'See if XML reponse is 200
		If intResponseCode = 200 Then 
			
			'Read in the data
			strNewVersionNumber = objXmlDoc.childNodes(1).childNodes(1).text
			strLicenseType = objXmlDoc.childNodes(1).childNodes(2).text
			strReleaseDate = objXmlDoc.childNodes(1).childNodes(3).text
			strReleaseAbout = objXmlDoc.childNodes(1).childNodes(4).text
			strNewsFeed = objXmlDoc.childNodes(1).childNodes(5).text
			
			If strLicenseType = "Free" Then
				Call addConfigurationItem(strFID, -1)
				Call addConfigurationItem(strFID2, -1)
				Call addConfigurationItem("Install_ID", "")
				strLicenseType = "Free Edition"
				
			ElseIf strLicenseType = "Brand-Free" Then
				Call addConfigurationItem(strFID, 0)
				Call addConfigurationItem(strFID2, 0)
				strLicenseType = "Premium Edition Brand-Free"
			
			ElseIf strLicenseType = "Branded" Then
				Call addConfigurationItem(strFID, -1)
				Call addConfigurationItem(strFID2, 0)
				strLicenseType = "Premium Edition (with Web Wiz Branding)"
				
			ElseIf strLicenseType = "WWH" Then
				Call addConfigurationItem(strFID, 0)
				Call addConfigurationItem(strFID2, 0)
				strLicenseType = "Premium Edition Brand-Free (Web Wiz Hosted Solution)"
			
			End If
			
			
		'Else give server error
		Else
			strUpdateServerError = objXmlDoc.childNodes(1).childNodes(1).text
			
		End If
	End If
      
        Set objXmlDoc = Nothing	
End If	

Call closeDatabase()
	

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta name="copyright" content="Copyright (C) 2001-2019 Web Wiz" />
<title>Check For Updates</title>

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
<h1>Check For Updates</h1>
 <a href="admin_menu.asp" target="_self">Return to the the Admin Control Panel Menu</a><br />
 <br />
 <table border="0" cellpadding="4" cellspacing="1" bordercolor="#000000" class="tableBorder">
  <tr>
   <td align="left" class="tableLedger">Update Sever Response </td>
  </tr>
  <tr>
   <td class="tableRow"><p>
    <%
   	
'If error give message
If strUpdateServerError <> "" Then
	
	Response.Write(strUpdateServerError)

'Else display the result
Else
	
	'If running latest release
	If strVersion = strNewVersionNumber Then 
		
		Response.Write("You are currently running the latest stable release.")
	
	Else	
		'If not running lastest release
		Response.Write("<strong>You are <u>NOT</u> currently running the latest stable release!</strong>" & _
			"<br />Please check with the software vendor for the latest release.")
		
	End If
%>
    <br /><br />
    <strong>Edition:</strong> <% = strLicenseType %><br />
    <strong>Running Version:</strong> <% = strVersion %><br />
    <strong>Database Backend:</strong> <% = strDatabaseType %><br />
    <br />
    <strong>Lastest Version:</strong> <% = strNewVersionNumber %><br />
    <strong>Release Date:</strong> <% = strReleaseDate %><br />
    <strong>Release Notes:</strong> <% = strReleaseAbout %><br />
    <br />
    <strong>News:</strong> <% = strNewsFeed %><%
    
	
End If

%>
     </td>
  </tr>
 </table>
 <br />
 <br />
 <br />
 <!-- #include file="includes/admin_footer_inc.asp" --><%

'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******

%>