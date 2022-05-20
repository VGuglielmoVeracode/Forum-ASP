<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="functions/functions_date_time_format.asp" -->
<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz Forums
'**  http://www.webwizforums.com
'**                                                              
'**  Copyright ©2001-2019 Web Wiz Ltd. All rights reserved.   
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


Response.ContentType = "text/html"


'Clean up
Call closeDatabase()



'Make sure this page is not cached
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"




Dim objXMLHTTP		'MS XML Object
Dim objRSSFeedItem	'XML Feed Items
Dim sarryRSSFeedItem	'RSS Feed array
Dim strHTML		'HTML Table results
Dim intFeedTimeToLive	'Time to live in minutes
Dim intRowColour	'Row colour number
Dim strNewsSubject
Dim strNewsArticle




'Time to live (how long the RSS Feed is cached in minutes)
'0 will reload immediately, but place more strain on the server if the page is called to often
intFeedTimeToLive = 5


'If this is x minutes or the feed is not in the web servers memory then grab the feed
If DateDiff("n", Application(strAppPrefix & "rssWebWizForumsUpdated"), Now()) >= intFeedTimeToLive OR Application(strAppPrefix & "rssWebWizForumsContent") = "" Then

   
	'Create MS XML object
	Set objXMLHTTP = Server.CreateObject("MSXML.DOMDocument")
	
	'Set the type of request HTTP Request
	objXMLHTTP.setProperty "ServerHTTPRequest", True
	
	'Disable Asyncronouse response
	objXMLHTTP.async = False
	
	'Load the Web Wiz Forums RSS Feed
	objXMLHTTP.Load(strWebWizNewsPadURL & "RSS_news_feed.asp")
	
	'If there is an error display a message
	If objXMLHTTP.parseError.errorCode <> 0 Then Response.Write " <strong>Error:</strong> " & objXMLHTTP.parseError.reason
	 
	'Create a new XML object containing the RSS Feed items
	Set objRSSFeedItem = objXMLHTTP.getElementsByTagName("item")
	
	'Loop through each of the XML RSS Feed items and place it in an HTML table
	For Each sarryRSSFeedItem In objRSSFeedItem
	
		'Web Wiz NewsPad RSS Feed Item childNodes
		'0 = title
		'1 = link
		'2 = description (arctile)
		'3 = pubDate
		'4 = guid (perminent link)
		'5 = WebWizNewsPad:pubDateISO (ISO international date)
		
		
		'Clean up input to prevent XXS hack
		strNewsSubject = formatInput(sarryRSSFeedItem.childNodes(0).text)

		'Remove HTML from article for subject link title
		strNewsArticle = removeHTML(sarryRSSFeedItem.childNodes(2).text, 150, true)

		'Clean up input to prevent XXS hack
		strNewsArticle = formatInput(strNewsArticle)
		
		'Calculate the row colour
		intRowColour = intRowColour + 1
	       
	       	'Create XHTML table rows
		strHTML = strHTML & " <tr "
		
		If (intRowColour MOD 2 = 0 ) Then strHTML = strHTML & "class=""evenTableRow""" Else strHTML = strHTML & "class=""oddTableRow""" 
			 	
		strHTML = strHTML & ">" & _
		vbCrLf & "  <td><a href=""" & sarryRSSFeedItem.childNodes(4).text & """ title=""" & strNewsArticle & """>" & strNewsSubject & "</a></td>" & _
	 	vbCrLf & "  <td>" &  DateFormat(sarryRSSFeedItem.childNodes(5).text) & "</td>" & _
		vbCrLf & " </tr>"
	Next
	
	'Release the objects
	Set objXMLHTTP = Nothing
	Set objRSSFeedItem = Nothing
	
	'Stick the whole lot in a application array to boost performance 
	Application.Lock
	Application(strAppPrefix & "rssWebWizForumsContent") = strHTML
	Application(strAppPrefix & "rssWebWizForumsUpdated") = Now()
	Application.UnLock
End If


%>
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
 <tr class="tableLedger">
  <td width="70%"><a href="<% = strWebWizNewsPadURL %>"><% = strTxtNewsBulletins %></a></td>
  <td width="30%"><% = strTxtPublished %></td>
 </tr><%

'Display the Web Wiz Forums Posts in an HTML table
Response.Write(Application(strAppPrefix & "rssWebWizForumsContent"))

%>
</table>