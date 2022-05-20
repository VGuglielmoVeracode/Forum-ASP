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



Dim lngPostID
Dim blnAlreadyThanked
Dim blnMemberCreatedPost
Dim lngPostThanks
Dim lngPostAuthorID
Dim strPostAuthorUsername
Dim lngAuthorThanked
Dim intThreadNo
Dim lngAuthorPoints

blnAlreadyThanked = False
blnMemberCreatedPost = False


'Get the forum ID
lngPostID = LngC(Request("PID"))
intThreadNo = LngC(Request("ID"))



'see if the user is logged in and rating enabled
If blnPostThanks = False OR blnGuest OR blnActiveMember = False OR bannedIP() Then
	
	'Clean up
	Call closeDatabase()
	
	'Display message to user
	Response.Write(Server.HTMLEncode(strTxtYouMustHaveAnActiveMemberAccount) & ".")
	
	Response.Flush
	Response.End
	
End If




'If this is a post back update
If lngPostID > 0 Then

	
	'Check the database to make sure they are not voting for a topic they have started
	strSQL = "SELECT " & strDbTable & "Thread.Author_ID, " & strDbTable & "Thread.Thanks," & strDbTable & "Author.Username, " & strDbTable & "Author.Thanked, " & strDbTable & "Author.Points " & _
	"FROM " & strDbTable & "Thread" & strDBNoLock & ", " & strDbTable & "Author" & strDBNoLock & " " & _
	"WHERE " & strDbTable & "Thread.Author_ID = " & strDbTable & "Author.Author_ID " & _
		"AND " & strDbTable & "Thread.Thread_ID = " & lngPostID & ";"
	
					
	'Query the database
	rsCommon.Open strSQL, adoCon
					
	'If a record is returbed read in who created the post
	If NOT rsCommon.EOF Then
		strPostAuthorUsername = rsCommon("Username")
		lngPostAuthorID	= CLng(rsCommon("Author_ID"))		
		If isNumeric(rsCommon("Thanks")) Then lngPostThanks = CLng(rsCommon("Thanks")) Else lngPostThanks = 0
		If isNumeric(rsCommon("Thanked")) Then lngAuthorThanked = CLng(rsCommon("Thanked")) Else lngAuthorThanked = 0
		If isNumeric(rsCommon("Points")) Then lngAuthorPoints = CLng(rsCommon("Points")) Else lngAuthorPoints = 0
	End If								
				
	'Close the recordset
	rsCommon.Close
	
	
	'Check to make sure the thanker is not trying to thank themselves
	If lngPostAuthorID = lngLoggedInUserID Then
		
		'Set the member cerated booleon to true
		blnMemberCreatedPost = True 
		
		'Display message to user
		Response.Write(Server.HTMLEncode(strTxtYouCanNotThankYourself) & ".")
	End If
	

	'If the user did not start the topic check that they have not already voted and save that they have voted
	If blnMemberCreatedPost = False Then
		
		'Check the database to see if the user has already thanked member for this post
		strSQL = "SELECT " & strDbTable & "ThreadThanks.* " & _
		"FROM " & strDbTable & "ThreadThanks" & strRowLock & " " & _
		"WHERE " & strDbTable & "ThreadThanks.Thread_ID = " & lngPostID & " AND " & strDbTable & "ThreadThanks.Author_ID = " & lngLoggedInUserID & ";"
	
		'Set the cursor type property of the record set to Forward Only
		rsCommon.CursorType = 0
			
		'Set the Lock Type for the records so that the record set is only locked when it is updated
		rsCommon.LockType = 3
					
		'Query the database
		rsCommon.Open strSQL, adoCon
					
		'If a record is returned then the user has voted so set blnAlreadyThanked to true
		If NOT rsCommon.EOF Then
			
			'Set boolen to true that this member has already said thanks for this post
			blnAlreadyThanked = True
						
			'Display message to user
			Response.Write("<img src=""" & strImagePath & "thanks." & strForumImageType & """ title=""" & Server.HTMLEncode(strTxtThanks) & " (" & lngPostThanks & ")"" alt=""" & Server.HTMLEncode(strTxtThanks) & " (" & lngPostThanks & ")"" style=""vertical-align: text-bottom;"" /> " & Server.HTMLEncode(strTxtThanks) & " (" & lngPostThanks & ")" & _
				vbCrLf & " <br />" & Server.HTMLEncode(strTxtYouHaveAlreadySaidThanksForThisPost) & ".")
			
					
					
		'Else the user has not thanked the user so save that they have to the database
		Else		
			'Use ADO to update database as we already have a query running
			rsCommon.AddNew
			rsCommon.Fields("Thread_ID") = lngPostID
			rsCommon.Fields("Author_ID") = lngLoggedInUserID
			rsCommon.Update
		End If				
					
					
		'Close the recordset
		rsCommon.Close
	End If
	
	


	'Update the number of thanks for the post
	If blnAlreadyThanked = False AND blnMemberCreatedPost = False Then

		'Increament number of times the post has had thanks
		lngPostThanks = lngPostThanks + 1
		
		'SQL to update number of times the post has had thanks
		strSQL = "UPDATE " & strDbTable & "Thread " & strRowLock & " " & _
		"SET " & strDbTable & "Thread.Thanks = " & lngPostThanks & " " & _
		"WHERE " & strDbTable & "Thread.Thread_ID = " & lngPostID & ";"	
		
		'Excute SQL
		adoCon.Execute(strSQL)
		
		
		
		'Increament number of times the member has been thanked
		lngAuthorThanked = lngAuthorThanked + 1
		lngAuthorPoints = lngAuthorPoints + intPointsThanked
		
		'SQL to update number of times the member has been thanked
		strSQL = "UPDATE " & strDbTable & "Author " & strRowLock & " " & _
		"SET " & strDbTable & "Author.Thanked = " & lngAuthorThanked & ", " & strDbTable & "Author.Points = " & lngAuthorPoints & " " & _
		"WHERE " & strDbTable & "Author.Author_ID = " & lngPostAuthorID & ";"	
		
		'Excute SQL
		adoCon.Execute(strSQL)
		
		'Display message to user that thanks has been updated
		Response.Write("<img src=""" & strImagePath & "thanks." & strForumImageType & """ title=""" & Server.HTMLEncode(strTxtThanks & " (" & lngPostThanks) & ")"" alt=""" & Server.HTMLEncode(strTxtThanks) & " (" & lngPostThanks & ")"" style=""vertical-align: text-bottom;"" /> " & Server.HTMLEncode(strTxtThanks) & " (" & lngPostThanks & ")" & _
			vbCrLf & " <br />" & Server.HTMLEncode(strPostAuthorUsername & " " & strTxtHasBeenThankedForTheirPost) & ".")
		
	End If
	
End If



'Clean up
Call closeDatabase()

%>