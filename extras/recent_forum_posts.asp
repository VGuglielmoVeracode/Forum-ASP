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

Dim objXMLHTTP		'MS XML Object
Dim objRSSFeedItem	'XML Feed Items
Dim sarryRSSFeedItem	'RSS Feed array
Dim strHTML		'HTML Table results
Dim strWebWizForumsURL	'Web Wiz Forums RSS Feed
Dim intTimeToLive	'Time to live in minutes


'Holds the URL to the Web Wiz Forums RSS Feed
strWebWizForumsURL = "http://forums.webwiz.net/RSS_topic_feed.asp"


'Time to live (how long the RSS Feed is cached in minutes)
'0 will reload immediately, but place more strain on the server if the page is called to often
intTimeToLive = 2

%>
<table class="tableBorder" cellspacing="0" cellpadding="0" width="100%">
 <tr class="tableLedger">
  <td><strong>Recent Forum Posts</strong></td>
 </tr><%

'If this is x minutes or the feed is not in the web servers memory then grab the feed
If DateDiff("n", Application("rssWebWizForumsUpdated"), Now()) >= intTimeToLive OR Application("rssWebWizForumsContent") = "" Then

   
	'Create MS XML object
	Set objXMLHTTP = Server.CreateObject("MSXML.DOMDocument")
	
	'Set the type of request HTTP Request
	objXMLHTTP.setProperty "ServerHTTPRequest", True
	
	'Disable Asyncronouse response
	objXMLHTTP.async = False
	
	'Load the Web Wiz Forums RSS Feed
	objXMLHTTP.Load(strWebWizForumsURL)
	
	'If there is an error display a message
	If objXMLHTTP.parseError.errorCode <> 0 Then Response.Write " <tr class=""tableRow""><td><strong>Error:</strong> " & objXMLHTTP.parseError.reason & "</td></tr>"
	 
	'Create a new XML object containing the RSS Feed items
	Set objRSSFeedItem = objXMLHTTP.getElementsByTagName("item")
	
	'Loop through each of the XML RSS Feed items and place it in an HTML table
	For Each sarryRSSFeedItem In objRSSFeedItem
	
		'Web Wiz Forums RSS Feed Item childNodes
		'0 = title
		'1 = link
		'2 = description (post)
		'3 = pubDate
		'4 = guid (perminent link)
	       
		strHTML = strHTML & " <tr class=""tableRow"">" & _
		vbCrLf & "  <td>" & _
		vbCrLf & "   <a href=""" & sarryRSSFeedItem.childNodes(4).text & """ title=""Posted: " & sarryRSSFeedItem.childNodes(3).text & """>" & sarryRSSFeedItem.childNodes(0).text & "</a>"
		
		'If you wish to display the entire post, uncomment the line below
		'strHTML = strHTML & vbCrLf & "   <br />" &  sarryRSSFeedItem.childNodes(2).text & "<br /><br />"
		
		strHTML = strHTML & vbCrLf & "  </td>" & _
		vbCrLf & " </tr>"
	Next
	
	'Release the objects
	Set objXMLHTTP = Nothing
	Set objRSSFeedItem = Nothing
	
	'Stick the whole lot in a application array to boost performance 
	Application.Lock
	Application("rssWebWizForumsContent") = strHTML
	Application("rssWebWizForumsUpdated") = Now()
	Application.UnLock
End If

'Display the Web Wiz Forums Posts in an HTML table
Response.Write(Application("rssWebWizForumsContent"))
%>
</table>