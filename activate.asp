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



'Set the response buffer to true as we maybe redirecting and setting a cookie
Response.Buffer = True 


'Dimension variables
Dim strUserCode			'Holds a code for the user
Dim lngUserID			'Holds the new users ID number
Dim blnActivated		'Set to true if the account is activated
Dim lngNewUserID		'Holds the ID of the user
Dim blnDbActive			'Holds if the user is active in db
Dim strDbUsercode		'Holds if the users db user code
Dim blnDbBanned			'Holds if the user is suspended

blnActivated = False
blnDbActive = False

'Read in the users ID from the query string
strUserCode = Trim(Mid(Request.QueryString("ID"), 1, 44))
If Request.QueryString("USD") <> "" AND isNumeric(Request.QueryString("USD")) Then lngNewUserID = LngC(Trim(Mid(Request.QueryString("USD"), 1, 6))) Else lngNewUserID = 0

'Make the usercode SQL safe
strUserCode = formatSQLInput(strUserCode)


'If there is a usercode activate it
If strUserCode <> "" AND blnMemberApprove = False Then
	
	'Intialise the strSQL variable with an SQL string to open a record set for the Author table
	strSQL = "SELECT " & strDbTable & "Author.* " & _
	"From " & strDbTable & "Author" & strRowLock & " " & _
	"WHERE " & strDbTable & "Author.Author_ID = " & lngNewUserID & ";"
	
	'Set the cursor type property of the record set to Forward Only
	rsCommon.CursorType = 0
	
	'Set the Lock Type for the records so that the record set is only locked when it is updated
	rsCommon.LockType = 3
	
	'Open the author table
	rsCommon.Open strSQL, adoCon
	
	'If these a record returned then check that the user code matches up before activation
	If NOT rsCommon.EOF Then
	
		'Read in details for this user from the database
		blnDbActive = CBool(rsCommon("Active"))
		strDbUsercode = rsCommon("User_code")
		blnDbBanned = CBool(rsCommon("Banned"))
		
		'See if the user is already active
		If blnDbActive Then
			
			'Set the activate boolean to true
			blnActivated = True
		
			
		'Else activate the user account
		ElseIf strUserCode = strDbUsercode AND blnDbBanned = False Then
		
			'Calculate a new code for the user
			strUserCode = userCode(rsCommon("Username"))
	
			'Update the database by activating the users account
			rsCommon.Fields("User_code") = strUserCode
			rsCommon.Fields("Active") = True
				
			'Update the database with the new user's details
			rsCommon.Update
				
			'Write a cookie with the User ID number so the user logged in throughout the forum
			Call setCookie("sLID", "UID", strUserCode, True)
			Call saveSessionItem("UID", strUserCode)
			
			'Set the activate boolean to true
			blnActivated = True
		End If
	End If
	
	'Release objects
	rsCommon.Close
End If


'If active users is enabled update the active users application array
If blnActiveUsers Then
	'Call active users function
	saryActiveUsers = activeUsers(strTxtActivateAccount, "", "", 0)
End If

'Reset Server Objects
Call closeDatabase()


'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtActivateAccount

%>  
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title><% = strTxtActivateAccount %></title>
<meta name="generator" content="Web Wiz Forums" />
<meta name="robots" content="noindex, follow" />

<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write("<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Forums(TM) ver. " & strVersion & "" & _
vbCrLf & "Info: http://www.webwizforums.com" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>

<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />		
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% = strTxtActivateAccount %></h1></td>
</tr>
</table>
<br /><%

'If the account is now active display a message
If blnActivated = True Then
 %>
<table cellspacing="1" cellpadding="3" class="tableBorder" align="center">
  <tr class="tableLedger">
   <td colspan="2"><% = strTxtActivateAccount %></td>
  </tr>
  <tr class="tableRow">
   <td><strong><% = strTxtYourForumMemIsNowActive %></strong>
    <br />
    <br /><% =  strTxtYouCanAccessCP %><br /><a href="member_control_panel.asp<% = strQsSID1 %>"><% =  strTxtMemberCPMenu %></a>
    <br />
    <br /><a href="default.asp<% = strQsSID1 %>"><% = strTxtReturnToDiscussionForum %></a>
  </td>
 </tr>
</table><%

'Else there is an error so show error table
Else

%>
<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
  <tr>
   <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtError %></strong></td>
  </tr>
  <tr>
    <td><%

	'If their account has been suspended and they are trying to reactivate it then tell them they can't
	If blnDbBanned Then
		
		Response.Write(strTxtForumMemberSuspended)
	
	'Theres been a problem so display an error message
	Else 
	   	Response.Write(strTxtErrorWithActvation & " " & strMainForumName & " " & " <a href=""mailto:" & strForumEmailAddress & """>" & strTxtForumAdministrator & "</a>.")
	End If 
	
	%></td>
  </tr>
</table><%
	
End If

%>
<br /><br />
<div align="center"><% 

'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnLCode = True Then
	
	If blnTextLinks = True Then 
		Response.Write("<span class=""text"" style=""font-size:10px"">Forum Software by <a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank""  style=""font-size:10px"">Web Wiz Forums&reg;</a> version " & strVersion & "</span>")
	Else
  		Response.Write("<a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank"" ><img src=""webwizforums_image.asp"" border=""0"" title=""Forum Software by Web Wiz Forums&reg; version " & strVersion& """ alt=""Forum Software by Web Wiz Forums&reg; version " & strVersion& """ /></a>")
	End If
	
	Response.Write("<br /><span class=""text"" style=""font-size:10px"">Copyright &copy;2001-2019 Web Wiz Ltd.</span>")
End If 
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******

'Display the process time
If blnShowProcessTime Then Response.Write "<span class=""smText""><br /><br />" & strTxtThisPageWasGeneratedIn & " " & FormatNumber(Timer() - dblStartTime, 3) & " " & strTxtSeconds & "</span>"
%>
</div>
<!-- #include file="includes/footer.asp" -->