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

      

'Read in the users details for the forum
intTopicPerPage	= IntC(Request.Form("topic"))	
intMaxPollChoices = IntC(Request.Form("pollChoice"))	
intThreadsPerPage = IntC(Request.Form("threads"))
intNumHotViews = IntC(Request.Form("hotViews"))
intNumHotReplies = IntC(Request.Form("hotReplies"))
blnShowEditUser = BoolC(Request.Form("edited"))
blnFlashFiles = BoolC(Request.Form("flash"))
blnYouTube = BoolC(Request.Form("YouTube"))
blnTopicIcon = BoolC(Request.Form("TopicIcons"))
strDefaultPostOrder = Request.Form("PostDir")
blnQuickReplyForm = BoolC(Request.Form("quickReply"))
blnCAPTCHAsecurityImages = BoolC(Request.Form("CAPTCHA"))
intEditedTimeDelay = IntC(Request.Form("editedTime"))
blnTopicRating = BoolC(Request.Form("topicRating"))
blnBoldNewTopics = BoolC(Request.Form("newTopicsBold"))
blnSignatures = BoolC(Request.Form("sigs"))
intEditPostTimeFrame = IntC(Request.Form("editTimeFrame"))
strAnswerPosts = Request.Form("AnswerPosts")
strAnswerPostsWording = Request.Form("AnswerPostsWording")
blnShareTopicLinks = BoolC(Request.Form("ShareTopic"))
blnPostThanks = BoolC(Request.Form("PostThanks"))
blnFacebookLike = BoolC(Request.Form("FacebookLikes"))
blnTwitterTweet = BoolC(Request.Form("TwitterTweets"))
strFacebookPageID = Request.Form("FacebookPageID")
strFacebookImage = Request.Form("FacebookImage")
blnGooglePlusOne = Request.Form("googlePlus1")
blnLocalLinksSameWindow = Request.Form("LocalLinksWindow")
blnCovertTextLinkstoHyperLink = Request.Form("textToHyperlink")
			

'If the user is changing tthe colours then update the database
If Request.Form("postBack") AND blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
		
	
	Call addConfigurationItem("Topics_per_page", intTopicPerPage)
	Call addConfigurationItem("Threads_per_page", intThreadsPerPage)
	Call addConfigurationItem("Hot_views", intNumHotViews)
	Call addConfigurationItem("Hot_replies", intNumHotReplies)
	Call addConfigurationItem("Show_edit", blnShowEditUser)
	Call addConfigurationItem("Flash", blnFlashFiles)
	Call addConfigurationItem("Vote_choices", intMaxPollChoices)
	Call addConfigurationItem("Topic_icon", blnTopicIcon)
	Call addConfigurationItem("Post_order", strDefaultPostOrder)
	Call addConfigurationItem("YouTube", blnYouTube)
	Call addConfigurationItem("Quick_reply", blnQuickReplyForm)
	Call addConfigurationItem("CAPTCHA", blnCAPTCHAsecurityImages)
	Call addConfigurationItem("Edited_by_delay", intEditedTimeDelay)
	Call addConfigurationItem("Topic_rating", blnTopicRating)
	Call addConfigurationItem("Topics_new_bold", blnBoldNewTopics)
	Call addConfigurationItem("Signatures", blnSignatures)
	Call addConfigurationItem("Edit_post_time_frame", intEditPostTimeFrame)
	Call addConfigurationItem("Answer_posts", strAnswerPosts)
	Call addConfigurationItem("Answer_wording", strAnswerPostsWording)
	Call addConfigurationItem("Share_topics_links", blnShareTopicLinks)
	Call addConfigurationItem("Post_thanks", blnPostThanks)
	Call addConfigurationItem("Facebook_likes", blnFacebookLike)
	Call addConfigurationItem("Twitter_tweet", blnTwitterTweet)
	Call addConfigurationItem("Facebook_page_ID", strFacebookPageID)
	Call addConfigurationItem("Facebook_image", strFacebookImage)
	Call addConfigurationItem("Google_plus_1", blnGooglePlusOne)
	Call addConfigurationItem("Local_links_window", blnLocalLinksSameWindow)
	Call addConfigurationItem("Text_Link_To_Hyperlink", blnCovertTextLinkstoHyperLink)
					
	'Update variables
	Application.Lock
	
	
	
	Application(strAppPrefix & "intTopicPerPage") = CInt(intTopicPerPage)
	Application(strAppPrefix & "intThreadsPerPage") = CInt(intThreadsPerPage)
	Application(strAppPrefix & "intNumHotViews") = CInt(intNumHotViews)
	Application(strAppPrefix & "intNumHotReplies") = CInt(intNumHotReplies)
	Application(strAppPrefix & "blnShowEditUser") = CBool(blnShowEditUser)
	Application(strAppPrefix & "blnFlashFiles") = CBool(blnFlashFiles)
	Application(strAppPrefix & "intMaxPollChoices") = intMaxPollChoices
	Application(strAppPrefix & "blnTopicIcon") = CBool(blnTopicIcon)
	Application(strAppPrefix & "strDefaultPostOrder") = strDefaultPostOrder
	Application(strAppPrefix & "blnYouTube") = CBool(blnYouTube)
	Application(strAppPrefix & "blnQuickReplyForm") = CBool(blnQuickReplyForm)
	Application(strAppPrefix & "blnCAPTCHAsecurityImages") = CBool(blnCAPTCHAsecurityImages)
	Application(strAppPrefix & "intEditedTimeDelay") = CInt(intEditedTimeDelay)
	Application(strAppPrefix & "blnTopicRating") = CBool(blnTopicRating)
	Application(strAppPrefix & "blnBoldNewTopics") = CBool(blnBoldNewTopics)
	Application(strAppPrefix & "blnSignatures") = CBool(blnSignatures)
	Application(strAppPrefix & "intEditPostTimeFrame") = CBool(intEditPostTimeFrame)
	Application(strAppPrefix & "strAnswerPosts") = strAnswerPosts
	Application(strAppPrefix & "strAnswerPostsWording") = strAnswerPostsWording
	Application(strAppPrefix & "blnShareTopicLinks") = CBool(blnShareTopicLinks)
	Application(strAppPrefix & "blnPostThanks") = CBool(blnPostThanks)
	Application(strAppPrefix & "blnFacebookLike") = CBool(blnFacebookLike)
	Application(strAppPrefix & "blnTwitterTweet") = CBool(blnTwitterTweet)
	Application(strAppPrefix & "strFacebookPageID") = strFacebookPageID
	Application(strAppPrefix & "strFacebookImage") = strFacebookImage
	Application(strAppPrefix & "blnGooglePlusOne") = CBool(blnGooglePlusOne)
	Application(strAppPrefix & "blnLocalLinksSameWindow") = CBool(blnLocalLinksSameWindow)
	Application(strAppPrefix & "blnCovertTextLinkstoHyperLink") = CBool(blnCovertTextLinkstoHyperLink)
	
	
	'Empty the application level variable so that the changes made are seen in the main forum
	Application(strAppPrefix & "blnConfigurationSet") = false
	
	Application.UnLock
End If






'Initialise the SQL variable with an SQL statement to get the configuration details from the database
strSQL = "SELECT " & strDbTable & "SetupOptions.Option_Item, " & strDbTable & "SetupOptions.Option_Value " & _
"FROM " & strDbTable & "SetupOptions" &  strDBNoLock & " " & _
"ORDER BY " & strDbTable & "SetupOptions.Option_Item ASC;"

	
'Query the database
rsCommon.Open strSQL, adoCon

'Read in the forum from the database
If NOT rsCommon.EOF Then
	
	'Place into an array for performance
	saryConfiguration = rsCommon.GetRows()

	
	'Read in the colour info from the database
	intTopicPerPage = CInt(getConfigurationItem("Topics_per_page", "numeric"))
	intThreadsPerPage = CInt(getConfigurationItem("Threads_per_page", "numeric"))
	intNumHotViews = CInt(getConfigurationItem("Hot_views", "numeric"))
	intNumHotReplies = CInt(getConfigurationItem("Hot_replies", "numeric"))
	blnShowEditUser = CBool(getConfigurationItem("Show_edit", "bool"))
	blnFlashFiles = CBool(getConfigurationItem("Flash", "bool"))
	intMaxPollChoices = CInt(getConfigurationItem("Vote_choices", "numeric"))
	blnTopicIcon = CBool(getConfigurationItem("Topic_icon", "bool"))
	strDefaultPostOrder = getConfigurationItem("Post_order", "string")
	blnYouTube = CBool(getConfigurationItem("YouTube", "bool"))
	blnQuickReplyForm = CBool(getConfigurationItem("Quick_reply", "bool"))
	blnCAPTCHAsecurityImages = CBool(getConfigurationItem("CAPTCHA", "bool"))
	intEditedTimeDelay = CInt(getConfigurationItem("Edited_by_delay", "numeric"))
	blnTopicRating = CBool(getConfigurationItem("Topic_rating", "bool"))
	blnBoldNewTopics = CBool(getConfigurationItem("Topics_new_bold", "bool"))
	blnSignatures = CBool(getConfigurationItem("Signatures", "bool"))
	intEditPostTimeFrame = CInt(getConfigurationItem("Edit_post_time_frame", "numeric"))
	strAnswerPosts = getConfigurationItem("Answer_posts", "string")
	strAnswerPostsWording = getConfigurationItem("Answer_wording", "string")
	blnShareTopicLinks = CBool(getConfigurationItem("Share_topics_links", "bool"))
	blnPostThanks = CBool(getConfigurationItem("Post_thanks", "bool"))
	blnFacebookLike = CBool(getConfigurationItem("Facebook_likes", "bool"))
	blnTwitterTweet = CBool(getConfigurationItem("Twitter_tweet", "bool"))
	strFacebookPageID = getConfigurationItem("Facebook_page_ID", "string")
	strFacebookImage = getConfigurationItem("Facebook_image", "string")
	blnGooglePlusOne = CBool(getConfigurationItem("Google_plus_1", "bool"))
	blnLocalLinksSameWindow = CBool(getConfigurationItem("Local_links_window", "bool"))
	blnCovertTextLinkstoHyperLink = CBool(getConfigurationItem("Text_Link_To_Hyperlink", "bool"))
	
End If


'Reset Server Objects
rsCommon.Close
Call closeDatabase()
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Topic and Post Settings</title>
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
<div align="center">
  <h1> Topic and Post Settings</h1> 
  <br />
    <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
    <br />
    From here you can configure general options of the forum to customise the functions and look of your forum.</p>
    <br />
    <br />
</div>
<form action="admin_post_topic_configure.asp<% = strQsSID1 %>" method="post" name="frmConfiguration" id="frmConfiguration" onsubmit="return CheckForm();">
  <table border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
     <td colspan="2" class="tableLedger"> Topic  Settings</td>
    </tr>
    <tr>
     <td class="tableRow" width="59%">Topics Per Page:<br />
       <span class="smText">This is the number of Topics shown on each page.</span></td>
     <td valign="top" class="tableRow"><select name="topic"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option <% If intTopicPerPage = 10 Then Response.Write("selected") %>>10</option>
       <option <% If intTopicPerPage = 11 Then Response.Write("selected") %>>11</option>
       <option <% If intTopicPerPage = 12 Then Response.Write("selected") %>>12</option>
       <option <% If intTopicPerPage = 13 Then Response.Write("selected") %>>13</option>
       <option <% If intTopicPerPage = 14 Then Response.Write("selected") %>>14</option>
       <option <% If intTopicPerPage = 15 Then Response.Write("selected") %>>15</option>
       <option <% If intTopicPerPage = 16 Then Response.Write("selected") %>>16</option>
       <option <% If intTopicPerPage = 17 Then Response.Write("selected") %>>17</option>
       <option <% If intTopicPerPage = 18 Then Response.Write("selected") %>>18</option>
       <option <% If intTopicPerPage = 19 Then Response.Write("selected") %>>19</option>
       <option <% If intTopicPerPage = 20 Then Response.Write("selected") %>>20</option>
       <option <% If intTopicPerPage = 21 Then Response.Write("selected") %>>21</option>
       <option <% If intTopicPerPage = 22 Then Response.Write("selected") %>>22</option>
       <option <% If intTopicPerPage = 23 Then Response.Write("selected") %>>23</option>
       <option <% If intTopicPerPage = 24 Then Response.Write("selected") %>>24</option>
       <option <% If intTopicPerPage = 25 Then Response.Write("selected") %>>25</option>
       <option <% If intTopicPerPage = 26 Then Response.Write("selected") %>>26</option>
       <option <% If intTopicPerPage = 28 Then Response.Write("selected") %>>28</option>
       <option <% If intTopicPerPage = 30 Then Response.Write("selected") %>>30</option>
       <option <% If intTopicPerPage = 35 Then Response.Write("selected") %>>35</option>
       <option <% If intTopicPerPage = 40 Then Response.Write("selected") %>>40</option>
       <option <% If intTopicPerPage = 45 Then Response.Write("selected") %>>45</option>
       <option <% If intTopicPerPage = 50 Then Response.Write("selected") %>>50</option>
       <option <% If intTopicPerPage = 75 Then Response.Write("selected") %>>75</option>
       <option <% If intTopicPerPage = 100 Then Response.Write("selected") %>>100</option>
      </select>     </td>
    </tr>
    <tr>
     <td class="tableRow">Hot Topic Number of Views:<br />
       <span class="smText">This is the number of times a Topic is viewed before it is shown as a Hot Topic.</span></td>
     <td valign="top" class="tableRow"><select name="hotViews"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option<% If intNumHotViews = 5 Then Response.Write(" selected") %>>5</option>
       <option<% If intNumHotViews = 10 Then Response.Write(" selected") %>>10</option>
       <option<% If intNumHotViews = 20 Then Response.Write(" selected") %>>20</option>
       <option<% If intNumHotViews = 30 Then Response.Write(" selected") %>>30</option>
       <option<% If intNumHotViews = 40 Then Response.Write(" selected") %>>40</option>
       <option<% If intNumHotViews = 50 Then Response.Write(" selected") %>>50</option>
       <option<% If intNumHotViews = 60 Then Response.Write(" selected") %>>60</option>
       <option<% If intNumHotViews = 70 Then Response.Write(" selected") %>>70</option>
       <option<% If intNumHotViews = 80 Then Response.Write(" selected") %>>80</option>
       <option<% If intNumHotViews = 90 Then Response.Write(" selected") %>>90</option>
       <option<% If intNumHotViews = 100 Then Response.Write(" selected") %>>100</option>
       <option<% If intNumHotViews = 110 Then Response.Write(" selected") %>>110</option>
       <option<% If intNumHotViews = 120 Then Response.Write(" selected") %>>120</option>
       <option<% If intNumHotViews = 130 Then Response.Write(" selected") %>>130</option>
       <option<% If intNumHotViews = 140 Then Response.Write(" selected") %>>140</option>
       <option<% If intNumHotViews = 150 Then Response.Write(" selected") %>>150</option>
       <option<% If intNumHotViews = 200 Then Response.Write(" selected") %>>200</option>
       <option<% If intNumHotViews = 250 Then Response.Write(" selected") %>>250</option>
       <option<% If intNumHotViews = 300 Then Response.Write(" selected") %>>300</option>
       <option<% If intNumHotViews = 400 Then Response.Write(" selected") %>>400</option>
       <option<% If intNumHotViews = 500 Then Response.Write(" selected") %>>500</option>
       <option<% If intNumHotViews = 999 Then Response.Write(" selected") %>>999</option>
      </select>     </td>
    </tr>
    <tr>
     <td class="tableRow">Hot Topic Number or Replies:<br />
       <span class="smText">This is the number of Replies a Topic must have to be shown as a Hot Topic.</span></td>
     <td valign="top" class="tableRow"><select name="hotReplies"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option<% If intNumHotReplies = 3 Then Response.Write(" selected") %>>3</option>
       <option<% If intNumHotReplies = 4 Then Response.Write(" selected") %>>4</option>
       <option<% If intNumHotReplies = 5 Then Response.Write(" selected") %>>5</option>
       <option<% If intNumHotReplies = 6 Then Response.Write(" selected") %>>6</option>
       <option<% If intNumHotReplies = 7 Then Response.Write(" selected") %>>7</option>
       <option<% If intNumHotReplies = 8 Then Response.Write(" selected") %>>8</option>
       <option<% If intNumHotReplies = 9 Then Response.Write(" selected") %>>9</option>
       <option<% If intNumHotReplies = 10 Then Response.Write(" selected") %>>10</option>
       <option<% If intNumHotReplies = 15 Then Response.Write(" selected") %>>15</option>
       <option<% If intNumHotReplies = 20 Then Response.Write(" selected") %>>20</option>
       <option<% If intNumHotReplies = 25 Then Response.Write(" selected") %>>25</option>
       <option<% If intNumHotReplies = 30 Then Response.Write(" selected") %>>30</option>
       <option<% If intNumHotReplies = 35 Then Response.Write(" selected") %>>35</option>
       <option<% If intNumHotReplies = 40 Then Response.Write(" selected") %>>40</option>
       <option<% If intNumHotReplies = 45 Then Response.Write(" selected") %>>45</option>
       <option<% If intNumHotReplies = 50 Then Response.Write(" selected") %>>50</option>
       <option<% If intNumHotReplies = 60 Then Response.Write(" selected") %>>60</option>
       <option<% If intNumHotReplies = 75 Then Response.Write(" selected") %>>75</option>
       <option<% If intNumHotReplies = 100 Then Response.Write(" selected") %>>100</option>
     </select></td>
    </tr>
    <tr>
     <td class="tableRow">Topic Icons:<br />
      <span class="smText">When enabled this allows Members to select an icon for the topic which is shown next to the subject line.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="TopicIcons" value="True" <% If blnTopicIcon = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="TopicIcons" value="False" <% If blnTopicIcon = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
     <td class="tableRow">Topic Rating System:<br />
      <span class="smText">When enabled this allows Members to rate Topics on a scale of 1 to 5, 5 being the best.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="topicRating" value="True" <% If blnTopicRating = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="topicRating" value="False" <% If blnTopicRating = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
     <tr>
     <td class="tableRow">Display New Unread Topics in Bold:<br />
      <span class="smText">When enabled this displays any topics the member has not read in bold.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="newTopicsBold" value="True" <% If blnBoldNewTopics = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="newTopicsBold" value="False" <% If blnBoldNewTopics = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
      <td colspan="2" class="tableLedger">Posts Settings</td>
    </tr>
    <tr>
     <td class="tableRow">Posts Per Page:<br />
       <span class="smText">This is the number of Posts shown on each page of a Topic.</span></td>
     <td valign="top" class="tableRow"><select name="threads"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option<% If intThreadsPerPage = 3 Then Response.Write(" selected") %>>3</option>
       <option<% If intThreadsPerPage = 4 Then Response.Write(" selected") %>>4</option>
       <option<% If intThreadsPerPage = 5 Then Response.Write(" selected") %>>5</option>
       <option<% If intThreadsPerPage = 6 Then Response.Write(" selected") %>>6</option>
       <option<% If intThreadsPerPage = 7 Then Response.Write(" selected") %>>7</option>
       <option<% If intThreadsPerPage = 8 Then Response.Write(" selected") %>>8</option>
       <option<% If intThreadsPerPage = 9 Then Response.Write(" selected") %>>9</option>
       <option<% If intThreadsPerPage = 10 Then Response.Write(" selected") %>>10</option>
       <option<% If intThreadsPerPage = 11 Then Response.Write(" selected") %>>11</option>
       <option<% If intThreadsPerPage = 12 Then Response.Write(" selected") %>>12</option>
       <option<% If intThreadsPerPage = 13 Then Response.Write(" selected") %>>13</option>
       <option<% If intThreadsPerPage = 14 Then Response.Write(" selected") %>>14</option>
       <option<% If intThreadsPerPage = 15 Then Response.Write(" selected") %>>15</option>
       <option<% If intThreadsPerPage = 16 Then Response.Write(" selected") %>>16</option>
       <option<% If intThreadsPerPage = 17 Then Response.Write(" selected") %>>17</option>
       <option<% If intThreadsPerPage = 18 Then Response.Write(" selected") %>>18</option>
       <option<% If intThreadsPerPage = 19 Then Response.Write(" selected") %>>19</option>
       <option<% If intThreadsPerPage = 20 Then Response.Write(" selected") %>>20</option>
       <option<% If intThreadsPerPage = 25 Then Response.Write(" selected") %>>25</option>
       <option<% If intThreadsPerPage = 30 Then Response.Write(" selected") %>>30</option>
       <option<% If intThreadsPerPage = 35 Then Response.Write(" selected") %>>35</option>
       <option<% If intThreadsPerPage = 40 Then Response.Write(" selected") %>>40</option>
       <option<% If intThreadsPerPage = 45 Then Response.Write(" selected") %>>45</option>
       <option<% If intThreadsPerPage = 50 Then Response.Write(" selected") %>>50</option>
       <option<% If intThreadsPerPage = 75 Then Response.Write(" selected") %>>75</option>
       <option<% If intThreadsPerPage = 100 Then Response.Write(" selected") %>>100</option>
       <option<% If intThreadsPerPage = 150 Then Response.Write(" selected") %>>150</option>
       <option<% If intThreadsPerPage = 200 Then Response.Write(" selected") %>>200</option>
       <option<% If intThreadsPerPage = 250 Then Response.Write(" selected") %>>250</option>
       <option<% If intThreadsPerPage = 300 Then Response.Write(" selected") %>>300</option>
       <option<% If intThreadsPerPage = 500 Then Response.Write(" selected") %>>500</option>
       <option<% If intThreadsPerPage = 999 Then Response.Write(" selected") %>>999</option>
      </select>     </td>
    </tr>
    <tr>
     <td class="tableRow">Post Display Order:<br />
       <span class="smText">This is the default order in which Posts are displayed within Topics.</span></td>
     <td valign="top" class="tableRow"><select name="PostDir"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option value="ASC"<% If strDefaultPostOrder = "ASC" Then Response.Write(" selected") %>>Oldest Post First</option>
       <option value="DESC"<% If strDefaultPostOrder = "DESC" Then Response.Write(" selected") %>>Newest Post First</option>
     </select></td>
    </tr>
    <tr>
     <td class="tableRow">Number of Poll Choices:<br />
       <span class="smText">This is the maximum number of choices allowed in a Forum Poll.</span></td>
     <td valign="top" class="tableRow"><select name="pollChoice" id="pollChoice"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option<% If intMaxPollChoices = 3 Then Response.Write(" selected") %>>3</option>
       <option<% If intMaxPollChoices = 4 Then Response.Write(" selected") %>>4</option>
       <option<% If intMaxPollChoices = 5 Then Response.Write(" selected") %>>5</option>
       <option<% If intMaxPollChoices = 6 Then Response.Write(" selected") %>>6</option>
       <option<% If intMaxPollChoices = 7 Then Response.Write(" selected") %>>7</option>
       <option<% If intMaxPollChoices = 8 Then Response.Write(" selected") %>>8</option>
       <option<% If intMaxPollChoices = 9 Then Response.Write(" selected") %>>9</option>
       <option<% If intMaxPollChoices = 10 Then Response.Write(" selected") %>>10</option>
       <option<% If intMaxPollChoices = 11 Then Response.Write(" selected") %>>11</option>
       <option<% If intMaxPollChoices = 12 Then Response.Write(" selected") %>>12</option>
       <option<% If intMaxPollChoices = 13 Then Response.Write(" selected") %>>13</option>
       <option<% If intMaxPollChoices = 14 Then Response.Write(" selected") %>>14</option>
       <option<% If intMaxPollChoices = 15 Then Response.Write(" selected") %>>15</option>
       <option<% If intMaxPollChoices = 16 Then Response.Write(" selected") %>>16</option>
       <option<% If intMaxPollChoices = 17 Then Response.Write(" selected") %>>17</option>
       <option<% If intMaxPollChoices = 18 Then Response.Write(" selected") %>>18</option>
       <option<% If intMaxPollChoices = 19 Then Response.Write(" selected") %>>19</option>
       <option<% If intMaxPollChoices = 20 Then Response.Write(" selected") %>>20</option>
       <option<% If intMaxPollChoices = 25 Then Response.Write(" selected") %>>25</option>
      </select>     </td>
    </tr>
    <tr>
     <td class="tableRow">Quick Reply Form:<br />
      <span class="smText">This displays the Quick Reply Post Form at the bottom Topic Pages.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" id="quickReply" name="quickReply" value="True" <% If blnQuickReplyForm = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" id="quickReply" name="quickReply" value="False" <% If blnQuickReplyForm = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
     <td class="tableRow">Post Editing Time Frame:<br />
       <span class="smText">This sets a time frame in which a member can Edit their Post. Once this time has expired the member can no-longer Edit their Post. <br />(Forum Admins and Moderators are not subject to this time frame and can Edit Posts at anytime)</span></td>
     <td valign="top" class="tableRow"><select name="editTimeFrame" id="editTimeFrame"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option<% If intEditPostTimeFrame = 0 Then Response.Write(" selected") %> value="0">Off</option>
       <option<% If intEditPostTimeFrame = 5 Then Response.Write(" selected") %> value="5">5 minutes</option>
       <option<% If intEditPostTimeFrame = 10 Then Response.Write(" selected") %> value="10">10 minutes</option>
       <option<% If intEditPostTimeFrame = 15 Then Response.Write(" selected") %> value="15">15 minutes</option>
       <option<% If intEditPostTimeFrame = 20 Then Response.Write(" selected") %> value="20">20 minutes</option>
       <option<% If intEditPostTimeFrame = 25 Then Response.Write(" selected") %> value="25">25 minutes</option>
       <option<% If intEditPostTimeFrame = 30 Then Response.Write(" selected") %> value="30">30 minutes</option>
       <option<% If intEditPostTimeFrame = 45 Then Response.Write(" selected") %> value="45">45 minutes</option>
       <option<% If intEditPostTimeFrame = 60 Then Response.Write(" selected") %> value="60">1 hour</option>
       <option<% If intEditPostTimeFrame = 120 Then Response.Write(" selected") %> value="120">2 hours</option>
       <option<% If intEditPostTimeFrame = 180 Then Response.Write(" selected") %> value="180">3 hours</option>
       <option<% If intEditPostTimeFrame = 360 Then Response.Write(" selected") %> value="360">6 hours</option>
       <option<% If intEditPostTimeFrame = 720 Then Response.Write(" selected") %> value="720">12 hours</option>
       <option<% If intEditPostTimeFrame = 1440 Then Response.Write(" selected") %> value="1440">1 Day</option>
       <option<% If intEditPostTimeFrame = 2880 Then Response.Write(" selected") %> value="2880">2 Days</option>
       <option<% If intEditPostTimeFrame = 4320 Then Response.Write(" selected") %> value="4320">3 Days</option>
       <option<% If intEditPostTimeFrame = 5760 Then Response.Write(" selected") %> value="5760">4 Days</option>
       <option<% If intEditPostTimeFrame = 7200 Then Response.Write(" selected") %> value="7200">5 Days</option>
       <option<% If intEditPostTimeFrame = 8640 Then Response.Write(" selected") %> value="8640">6 Days</option>
       <option<% If intEditPostTimeFrame = 10080 Then Response.Write(" selected") %> value="10080">1 Week</option>
       <option<% If intEditPostTimeFrame = 20160 Then Response.Write(" selected") %> value="20160">2 Weeks</option>
      </select>
     </td>
    </tr>
   
    <tr>
     <td class="tableRow">Signatures:<br />
       <span class="smText">When enabled this allows users to attach signatures to their posts. You can also disable signatures for individual Member Groups through the Group Admin section, if this option is first enabled.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="sigs" value="True" <% If blnSignatures = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="sigs" value="False" <% If blnSignatures = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td></tr>
   
     
    <tr>
     <td class="tableRow">Display Edited Member Name:<br />
       <span class="smText">When enabled this displays on the bottom of edited posts  the name of member and the date and time the post was edited.</span><br />     </td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="edited" value="True" <% If blnShowEditUser = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="edited" value="False" <% If blnShowEditUser = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
     <td class="tableRow">Display Edited Member Name Time Delay:<br />
       <span class="smText">This sets a time delay, so if a post is edited within this time after it was initially posted it will not display the  name of member and the date and time the post was edited</span>.</td>
      <td valign="top" class="tableRow"><select name="editedTime" id="editedTime"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option<% If intEditedTimeDelay = 0 Then Response.Write(" selected") %> value="0">Off</option>
       <option<% If intEditedTimeDelay = 1 Then Response.Write(" selected") %> value="1">1 minute</option>
       <option<% If intEditedTimeDelay = 2 Then Response.Write(" selected") %> value="2">2 minutes</option>
       <option<% If intEditedTimeDelay = 3 Then Response.Write(" selected") %> value="3">3 minutes</option>
       <option<% If intEditedTimeDelay = 4 Then Response.Write(" selected") %> value="4">4 minutes</option>
       <option<% If intEditedTimeDelay = 5 Then Response.Write(" selected") %> value="5">5 minutes</option>
       <option<% If intEditedTimeDelay = 6 Then Response.Write(" selected") %> value="6">6 minutes</option>
       <option<% If intEditedTimeDelay = 7 Then Response.Write(" selected") %> value="7">7 minutes</option>
       <option<% If intEditedTimeDelay = 8 Then Response.Write(" selected") %> value="8">8 minutes</option>
       <option<% If intEditedTimeDelay = 9 Then Response.Write(" selected") %> value="9">9 minutes</option>
       <option<% If intEditedTimeDelay = 10 Then Response.Write(" selected") %> value="10">10 minutes</option>
       <option<% If intEditedTimeDelay = 15 Then Response.Write(" selected") %> value="15">15 minutes</option>
       <option<% If intEditedTimeDelay = 20 Then Response.Write(" selected") %> value="20">20 minutes</option>
       <option<% If intEditedTimeDelay = 30 Then Response.Write(" selected") %> value="30">30 minutes</option>
       <option<% If intEditedTimeDelay = 60 Then Response.Write(" selected") %> value="60">1 hour</option>
      </select>
     </td>
    </tr>
    <tr>
     <td class="tableRow">Set Answer/Resolution Posts:<br />
       <span class="smText">This option is useful for Support and Question Forums as it allows Posts to be set as 'Answers' to the initial question posted by the Topic Starter. Posts set as the 'Answer' are highlighted and set as the Second Post within the Topic as the Answer or Resolution to the Topic Starters Question.</span></td>
     <td valign="top" class="tableRow">
      <select name="AnswerPosts" id="AnswerPosts"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option<% If strAnswerPosts = "Off" Then Response.Write(" selected") %> value="Off">Off</option>
       <option<% If strAnswerPosts = "admin" Then Response.Write(" selected") %> value="admin">Admins can set Posts as Answers</option>
       <option<% If strAnswerPosts = "admin_mods" Then Response.Write(" selected") %> value="admin_mods">Admins and Moderators can set Posts as Answers</option>
       <option<% If strAnswerPosts = "all_members" Then Response.Write(" selected") %> value="all_members">Admins, Moderators and Forum Members can set Posts as Answers</option>
      </select>
      </td>
    </tr>
    <td class="tableRow">Set Answer/Resolution Posts Wording:<br />
       <span class="smText">This option is for the Wording displayed in the forum system for Answer/Resolution Posts.</span></td>
     <td valign="top" class="tableRow">
      <select name="AnswerPostsWording" id="AnswerPostsWording"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option<% If strAnswerPostsWording = strTxtAnswer Then Response.Write(" selected") %> value="<% = strTxtAnswer %>"><% = strTxtAnswer %></option>
       <option<% If strAnswerPostsWording = strTxtResolution Then Response.Write(" selected") %> value="<% = strTxtResolution %>"><% = strTxtResolution %></option>
       <option<% If strAnswerPostsWording = strTxtOfficialResponse Then Response.Write(" selected") %> value="<% = strTxtOfficialResponse %>"><% = strTxtOfficialResponse %></option>
      </select>
     </td>
    </tr>
    <tr>
     <td class="tableRow">Thank Members:<br />
       <span class="smText">When enabled this then members can thank other members for Posts they have add to Forums.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="PostThanks" value="True" <% If blnPostThanks = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="PostThanks" value="False" <% If blnPostThanks = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td></tr>
    
     <tr>
     <td class="tableRow">Convert Text Links to Hyperlinks:<br />
       <span class="smText">When enabled hyperlinks that are in text format and not clickable are converted to clickable hyperlinks if hyperlinks are enabled for the Posters Group.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="textToHyperlink" value="True" <% If blnCovertTextLinkstoHyperLink = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="textToHyperlink" value="False" <% If blnCovertTextLinkstoHyperLink = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
     </tr>
    
    <tr>
     <td class="tableRow">Open Local Links in the same Browser Window:<br />
       <span class="smText">When enabled Local Links to your Website or Forum will open in the same Browser Window and not a New Browser Window.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="LocalLinksWindow" value="True" <% If blnLocalLinksSameWindow = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="LocalLinksWindow" value="False" <% If blnLocalLinksSameWindow = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
     </tr>
    
    <tr>
     <td class="tableRow">Adobe Flash:<br />
       <span class="smText">When enabled this then users will be able to display Flash content in their posts and signatures using Forum BBcode [FLASH]file name here[/FLASH]</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="flash" value="True" <% If blnFlashFiles = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="flash" value="False" <% If blnFlashFiles = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td></tr>
    
    <td class="tableRow">YouTube Movies:<br />
       <span class="smText">When enabled this then users will be able to display YouTube movies in their  posts and signatures using Forum BBcode [TUBE]file name here[/TUBE]</span>
     </td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="YouTube" value="True" <% If blnYouTube = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="YouTube" value="False" <% If blnYouTube = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
     </td>
    </tr>
     
    <td class="tableRow">Topic Sharing and Social Bookmarks:<br />
       <span class="smText">When enabled this will display a 'Share' button at the bottom of Post Pages with links to Share and Bookmark the Topic with popular Social Networking websites such as Facebook, Reddit, Dig, StumbleUpon, Buzz, Yahoo and more.</span>
     </td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="ShareTopic" value="True" <% If blnShareTopicLinks = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="ShareTopic" value="False" <% If blnShareTopicLinks = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
     </td>
    </tr>
    
     <td class="tableRow">Twitter Tweet Button:<br />
       <span class="smText">When enabled this will display a Twitter 'Tweet' button on pages displaying Posts.</span>
     </td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="TwitterTweets" value="True" <% If blnTwitterTweet = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="TwitterTweets" value="False" <% If blnTwitterTweet = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
     </td>
    </tr>
    
    <td class="tableRow">Google +1 Button:<br />
       <span class="smText">When enabled this will display a Google +1 button on pages displaying Posts.</span>
     </td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="googlePlus1" value="True" <% If blnGooglePlusOne = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="googlePlus1" value="False" <% If blnGooglePlusOne = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
     </td>
    </tr>
    
     <td class="tableRow">Facebook Like Button:<br />
       <span class="smText">When enabled this will display a Facebook 'Like' button on pages displaying Posts.</span>
     </td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="FacebookLikes" value="True" <% If blnFacebookLike = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="FacebookLikes" value="False" <% If blnFacebookLike = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
     </td>
    </tr>
    
    <td class="tableRow">Facebook Like Image:<br />
       <span class="smText">This is the image that you wish to be used when the Post is Liked or Shared on Facebook.</span>
     </td>
     <td valign="top" class="tableRow"><input name="FacebookImage" type="text" id="siteURL" value="<% = strFacebookImage %>" size="30" maxlength="70"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    
    <td class="tableRow">Facebook Fan Page ID:<br />
       <span class="smText">If you have a Fan or Company Facebook Page this is the Page ID to allow any Likes or Shares to be associated with your Facebook Page.</span>
     </td>
     <td valign="top" class="tableRow"><input name="FacebookPageID" type="text" id="siteURL" value="<% = strFacebookPageID %>" size="30" maxlength="70"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    
   
    
    <tr>
     <td class="tableRow"><a href="http://www.webwizcaptcha.com" target="_blank">Web Wiz CAPTCHA</a> for Guest Posting:<br />
       <span class="smText">This displays a security image when a Guest Posts. This prevents spamming from remote form submission which could flood your forum with unwanted spam posts if Guest Post in allowed.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="CAPTCHA" value="True" <% If blnCAPTCHAsecurityImages = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" value="False" <% If blnCAPTCHAsecurityImages = False Then Response.Write "checked" %> name="CAPTCHA"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr align="center">
      <td colspan="2" class="tableBottomRow">
          <input type="hidden" name="postBack" value="true" />
          <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
          <input type="submit" name="Submit" value="Update Topic and Post Settings" />
          <input type="reset" name="Reset" value="Reset Form" />      </td>
    </tr>
  </table>
</form>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
