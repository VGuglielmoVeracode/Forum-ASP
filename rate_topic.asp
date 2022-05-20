<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
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




'Set the response buffer to true as we maybe redirecting
Response.Buffer = True


'Make sure this page is not cached
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"



Dim lngTopicID
Dim intRating 
Dim blnAlreadyVoted
Dim blnTopicStarter
Dim dblTopicRating		'Holds the rating for a topic
Dim lngTopicRatingVotes		'Number of votes a topic receives
Dim lngTopicRatingTotal

blnAlreadyVoted = False
blnTopicStarter = False


'Get the forum ID
lngTopicID = LngC(Request("TID"))




'see if the user is logged in and rating enabled
If blnTopicRating = False OR blnGuest OR blnActiveMember = False OR bannedIP() Then
	
	'Clean up
	Call closeDatabase()
	
	Response.Redirect("default.asp")
	
End If


'If this is a post back take the rating
If Request.QueryString("postBack") Then

	'Read in the value
	intRating = IntC(Request.QueryString("rating"))
	
	
	
	'Check the database to make sure they are not voting for a topic they have started
	strSQL = "SELECT " & strDbTable & "Thread.Author_ID " & _
	"FROM " & strDbTable & "Thread" & strDBNoLock & ",  " & strDbTable & "Topic" & strDBNoLock & " " & _
	"WHERE " & strDbTable & "Topic.Start_Thread_ID = " & strDbTable & "Thread.Thread_ID " & _
		"AND " & strDbTable & "Topic.Topic_ID = " & lngTopicID & " " & _
		"AND " & strDbTable & "Thread.Author_ID = " & lngLoggedInUserID & ";"
					
	'Query the database
	rsCommon.Open strSQL, adoCon
					
	'If a record is returned then the user has started this topic
	If NOT rsCommon.EOF Then
						
		blnTopicStarter = True
	End If				
					
				
	'Close the recordset
	rsCommon.Close
	

	'If the user did not start the topic check that they have not already voted and save that they have voted
	If blnTopicStarter = False Then
		'Check the database to see if the user has voted
		strSQL = "SELECT " & strDbTable & "TopicRatingVote.* " & _
		"FROM " & strDbTable & "TopicRatingVote" & strRowLock & " " & _
		"WHERE " & strDbTable & "TopicRatingVote.Topic_ID = " & lngTopicID & " AND " & strDbTable & "TopicRatingVote.Author_ID = " & lngLoggedInUserID & ";"
	
		'Set the cursor type property of the record set to Forward Only
		rsCommon.CursorType = 0
			
		'Set the Lock Type for the records so that the record set is only locked when it is updated
		rsCommon.LockType = 3
					
		'Query the database
		rsCommon.Open strSQL, adoCon
					
		'If a record is returned then the user has voted so set blnAlreadyVoted to true
		If NOT rsCommon.EOF Then
						
			blnAlreadyVoted = True
					
					
		'Else the user has not voted so save there ID to database
		ElseIf intRating > 0 Then			
			'Use ADO to update database as we already have a query running
			rsCommon.AddNew
			rsCommon.Fields("Topic_ID") = lngTopicID
			rsCommon.Fields("Author_ID") = lngLoggedInUserID
			rsCommon.Update
		End If				
					
					
		'Close the recordset
		rsCommon.Close
	End If
	
	


	'Save the voters choice

	'Initlise the SQL query
	strSQL = "SELECT " & strDbTable & "Topic.Rating, " & strDbTable & "Topic.Rating_Votes, " & strDbTable & "Topic.Rating_Total " & _
	"FROM " & strDbTable & "Topic" & strRowLock & " " & _
	"WHERE " & strDbTable & "Topic.Topic_ID = " & lngTopicID & ";"

	'Query the database
	rsCommon.Open strSQL, adoCon

	'If a record is returned calculate the total for it
	If NOT rsCommon.EOF Then
		
		'Get data from db
		If isNumeric(rsCommon("Rating")) Then dblTopicRating = CDbl(rsCommon("Rating"))	Else dblTopicRating = 0
		If isNumeric(rsCommon("Rating_Votes")) Then lngTopicRatingVotes = CLng(rsCommon("Rating_Votes")) Else lngTopicRatingVotes = 0		
		If isNumeric(rsCommon("Rating_Total")) Then lngTopicRatingTotal = CLng(rsCommon("Rating_Total")) Else lngTopicRatingTotal = 0
		
		'Calculate total now (it is then calculated again if it is an update)
		If lngTopicRatingVotes > 0 Then dblTopicRating = lngTopicRatingTotal / lngTopicRatingVotes
		
		'If the already voted boolean is not set then save the vote
		If blnAlreadyVoted = False AND blnTopicStarter = False AND intRating > 0 Then

			'Calculate totals
			lngTopicRatingTotal = lngTopicRatingTotal + intRating
			lngTopicRatingVotes = lngTopicRatingVotes + 1
			
			dblTopicRating = lngTopicRatingTotal / lngTopicRatingVotes
			
			'Update the database 
			'(do not use ADO Update as mySQL sets all fields that have 0 to this rating)
			strSQL = "UPDATE " & strDbTable & "Topic " & _
			"SET " & strDbTable & "Topic.Rating = " & dblTopicRating & ", " & _
			strDbTable & "Topic.Rating_Votes = " & lngTopicRatingVotes & ", " & _
			strDbTable & "Topic.Rating_Total = " & lngTopicRatingTotal & " " & _
			"WHERE " & strDbTable & "Topic.Topic_ID = " & lngTopicID & ";"
				
			'Update db
			adoCon.Execute(strSQL)
		End If
	End If

	'Close the recordset
	rsCommon.Close
	
	
	'If a rating has been submited show the user the result
	If intRating > 0  Then
		Response.Write(" <div style=""text-align:center;height:111px;"">" & _
		VbCrLf & "   <br style=""line-height: 8px;""/>")
		
		'Display message to user	
		If blnAlreadyVoted Then
			Response.Write(Server.HTMLEncode(strTxtYouHaveAlreadyRatedThisTopic))
		ElseIf blnTopicStarter Then
			Response.Write(Server.HTMLEncode(strTxtYouCanNotRateATopicYouStarted))
		Else
			Response.Write(Server.HTMLEncode(strTxtThankYouForRatingThisTopic))
		End If
		
		'Display vote amount
		Response.Write("<br /><br />" & Server.HTMLEncode(strTxtTopicRating) & ": <img src=""" & strImagePath & Mid(CStr(dblTopicRating + 0.5), 1, 1) & "_star_topic_rating." & strForumImageType & """ alt=""" & strTxtTopicRating & ": " & lngTopicRatingVotes & " " & strTxtVotes & ", " &  strTxtAverage & " " & FormatNumber(dblTopicRating, 2) & """ title=""" & strTxtTopicRating & ": " & lngTopicRatingVotes & " " & strTxtVotes & ", " &  strTxtAverage & " " & FormatNumber(dblTopicRating, 2) & """ style=""vertical-align: text-bottom;"" />" & _
		"<br />" & lngTopicRatingVotes & " " & Server.HTMLEncode(strTxtVotes) & ", " &  Server.HTMLEncode(strTxtAverage) & " " & FormatNumber(dblTopicRating, 2))
		
		Response.Write("<br /><br />" & _
		VbCrLf & " </div>")
		
		'Flush
		Response.Flush
		Response.End
	End If
	
	
	
End If



'Initlise the SQL query
strSQL = "SELECT " & strDbTable & "Topic.Rating, " & strDbTable & "Topic.Rating_Votes, " & strDbTable & "Topic.Rating_Total " & _
"FROM " & strDbTable & "Topic" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Topic.Topic_ID = " & lngTopicID & ";"

'Query the database
rsCommon.Open strSQL, adoCon

'If a record is returned calculate the total for it
If NOT rsCommon.EOF Then
		
	'Get data from db
	If isNumeric(rsCommon("Rating")) Then dblTopicRating = CDbl(rsCommon("Rating"))	Else dblTopicRating = 0
	If isNumeric(rsCommon("Rating_Votes")) Then lngTopicRatingVotes = CLng(rsCommon("Rating_Votes")) Else lngTopicRatingVotes = 0		
	If isNumeric(rsCommon("Rating_Total")) Then lngTopicRatingTotal = CLng(rsCommon("Rating_Total")) Else lngTopicRatingTotal = 0
		
	'Calculate total now (it is then calculated again if it is an update)
	If lngTopicRatingVotes > 0 Then dblTopicRating = lngTopicRatingTotal / lngTopicRatingVotes
End If
	
rsCommon.Close


'Clean up
Call closeDatabase()

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Rate Topic</title>
<meta name="generator" content="Web Wiz Forums" />
<meta name="robots" content="noindex, follow" />
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="includes/default_javascript_v9.js" type="text/javascript"></script>
<script language="javascript" type="text/javascript">
function ratingStars(stars, action){	
	for (i=0; i<=4; i++){
		var rateText = document.getElementById('ratingText');
		if (action == 'over'){
			if (stars <= i){
				document.getElementById('rateStars' + i).src = '<% = strImagePath %>rating_select_off.<% = strForumImageType %>';
			}else{
				document.getElementById('rateStars' + i).src = '<% = strImagePath %>rating_select_on.<% = strForumImageType %>';	
			}
			switch(stars){
				case 1:	rateText.innerHTML = '<% = Server.HTMLEncode(strTxtTerrible) %>'; break;
				case 2:	rateText.innerHTML = '<% = Server.HTMLEncode(strTxtPoor) %>'; break;
				case 3:	rateText.innerHTML = '<% = Server.HTMLEncode(strTxtAverage) %>'; break;
				case 4:	rateText.innerHTML = '<% = Server.HTMLEncode(strTxtGood) %>'; break;
				case 5:	rateText.innerHTML = '<% = Server.HTMLEncode(strTxtExcellent) %>'; break;
			}
		}else{
			
			document.getElementById('rateStars' + i).src = '<% = strImagePath %>rating_select_off.<% = strForumImageType %>';
			rateText.innerHTML = '&nbsp;';
		}
	}
}
</script>
</head>
<body class="dropDownTopicRating" style="border-width: 0px;visibility: visible;margin:4px;width:180px;">
<span id="ajaxRatingResult">
 <div style="text-align:center;">
  <span id="ratingText">&nbsp;</span>
  <br />
  <img id="rateStars0" onmouseover="ratingStars(1, 'over');" onmouseout="ratingStars(0, 'out');" src="<% = strImagePath %>rating_select_off.<% = strForumImageType %>" title="<% = strTxtRateThisTopicAs & " " & strTxtTerrible %>" alt="<% = strTxtRateThisTopicAs & " " & strTxtTerrible %>" style="cursor: pointer;" onclick="getAjaxData('rate_topic.asp?TID=<% = lngTopicID %>&postBack=True&rating=1<% = strQsSID2 %>', 'ajaxRatingResult');" />
  <img id="rateStars1" onmouseover="ratingStars(2, 'over');" onmouseout="ratingStars(0, 'out');" src="<% = strImagePath %>rating_select_off.<% = strForumImageType %>" title="<% = strTxtRateThisTopicAs & " " & strTxtPoor %>" alt="<% = strTxtRateThisTopicAs & " " & strTxtPoor %>" style="cursor: pointer;" onclick="getAjaxData('rate_topic.asp?TID=<% = lngTopicID %>&postBack=True&rating=2<% = strQsSID2 %>', 'ajaxRatingResult');" />
  <img id="rateStars2" onmouseover="ratingStars(3, 'over');" onmouseout="ratingStars(0, 'out');" src="<% = strImagePath %>rating_select_off.<% = strForumImageType %>" title="<% = strTxtRateThisTopicAs & " " & strTxtAverage %>" alt="<% = strTxtRateThisTopicAs & " " & strTxtAverage %>" style="cursor: pointer;" onclick="getAjaxData('rate_topic.asp?TID=<% = lngTopicID %>&postBack=True&rating=3<% = strQsSID2 %>', 'ajaxRatingResult');" />
  <img id="rateStars3" onmouseover="ratingStars(4, 'over');" onmouseout="ratingStars(0, 'out');" src="<% = strImagePath %>rating_select_off.<% = strForumImageType %>" title="<% = strTxtRateThisTopicAs & " " & strTxtGood %>" alt="<% = strTxtRateThisTopicAs & " " & strTxtGood %>" style="cursor: pointer;" onclick="getAjaxData('rate_topic.asp?TID=<% = lngTopicID %>&postBack=True&rating=4<% = strQsSID2 %>', 'ajaxRatingResult');" />
  <img id="rateStars4" onmouseover="ratingStars(5, 'over');" onmouseout="ratingStars(0, 'out');" src="<% = strImagePath %>rating_select_off.<% = strForumImageType %>" title="<% = strTxtRateThisTopicAs & " " & strTxtExcellent %>" alt="<% = strTxtRateThisTopicAs & " " & strTxtExcellent %>" style="cursor: pointer;" onclick="getAjaxData('rate_topic.asp?TID=<% = lngTopicID %>&postBack=True&rating=5<% = strQsSID2 %>', 'ajaxRatingResult');" /><%
     	
     	'Display vote amount
	Response.Write(VbCrLf & "     	<br /><br />" & Server.HTMLEncode(strTxtTopicRating) & ": ")
	If strClientBrowserVersion <> "MSIE6-" Then
		Response.Write("<img src=""" & strImagePath & Mid(CStr(dblTopicRating + 0.5), 1, 1) & "_star_topic_rating." & strForumImageType & """ alt=""" & strTxtTopicRating & ": " & lngTopicRatingVotes & " " & strTxtVotes & ", " &  strTxtAverage & " " & FormatNumber(dblTopicRating, 2) & """ title=""" & strTxtTopicRating & ": " & lngTopicRatingVotes & " " & strTxtVotes & ", " &  strTxtAverage & " " & FormatNumber(dblTopicRating, 2) & """ style=""vertical-align: text-bottom;"" />")
	End If
	Response.Write("<br />" & Server.HTMLEncode(lngTopicRatingVotes) & " " & strTxtVotes & ", " &  Server.HTMLEncode(strTxtAverage) & " " & FormatNumber(dblTopicRating, 2))
		
     	%>
 </div>
</table>
</span>
</body>
</html>