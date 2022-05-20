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





Dim intNotifiedNoOfPMs
Dim intNumOfNewPM

'If the user is logged in and there account is active display if they have private messages
If intGroupID <> 2 AND blnActiveMember AND blnPrivateMessages Then

	'Display the number of new pm's
	If intNoOfPms > 0 Then
		Response.Write("&nbsp;&nbsp;<img src=""" & strImagePath & "new_private_message." & strForumImageType & """ title=""" & intNoOfPms & " " & strTxtNewMessages & """ alt=""" & intNoOfPms & " " & strTxtNewMessages & """ /> <a href=""pm_welcome.asp" & strQsSID1 & """><strong>" & intNoOfPms & "</strong> " & strTxtNewMessages & "</a>")
	Else
		Response.Write("&nbsp;&nbsp;<img src=""" & strImagePath & "private_message." & strForumImageType & """ title=""0 " & strTxtNewMessages & """ alt=""0 " & strTxtNewMessages & """ /> <a href=""pm_welcome.asp" & strQsSID1 & """>0 " & strTxtNewMessages & "</a>")
	End If

	'Get the number of PM's the user has been notified of from cookie
	If isNumeric(getSessionItem("PMN")) Then intNotifiedNoOfPMs = LngC(getSessionItem("PMN")) Else intNotifiedNoOfPMs = 0	
	
	'If the number of un-read PM's is higher then the user has been notified of, they have a new PM so tell them
	If intNoOfPms > intNotifiedNoOfPMs Then
			
		Call saveSessionItem("PMN", intNoOfPms)
			
		'Display the alert
		Response.Write("<script language=""JavaScript""><!-- " & _
		vbCrLf & vbCrLf & "//Display pop for new private message" & _
		vbCrLf & "checkPrivateMsg = confirm('" & strTxtYouHave & " " & intNoOfPms & " " & strTxtNewPMsClickToGoNowToPM & "')" & _
		vbCrLf & "if (checkPrivateMsg == true) {" & _
		vbCrLf & "	window.location='pm_inbox.asp" & strQsSID1 & "'" & _
		vbCrLf & "}" & _
		vbCrLf & "// --></script>")
	
	'If PM inbox is full tell the member so (no alert for admin acocunts which are unlimited)
	ElseIf intNoOfInboxPms >= intPmInbox AND BoolC(getSessionItem("PMI")) = False AND intGroupID > 1 Then
		
		Call saveSessionItem("PMI", 1)
		
		'Display the alert
		Response.Write("<script language=""JavaScript""><!-- " & _
		vbCrLf & vbCrLf & "//Display pop for full private message inbox" & _
		vbCrLf & "checkPrivateMsg = confirm('" & strTxtYourPmInboxIsFullPleaseDeleteOldPMs & "')" & _
		vbCrLf & "if (checkPrivateMsg == true) {" & _
		vbCrLf & "	window.location='pm_inbox.asp" & strQsSID1 & "'" & _
		vbCrLf & "}" & _
		vbCrLf & "// --></script>")
	
	End If
End If
%>