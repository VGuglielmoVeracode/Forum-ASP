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



'Dimension variables
Dim strMemberName	'Holds the member name to lookup
Dim blnMemNotFound	'Holds the error code if user not found
Dim lngMemberID		'Holds the ID number of the member




'If this is a postback check for the user exsisting in the db before redirecting
If Request.Form("postBack") Then
	
	'Initliase varaibles
	blnMemNotFound = false
	
	'Read in the members name to lookup
	strMemberName = Request.Form("member")
	
	'Get rid of milisous code
	strMemberName = formatSQLInput(strMemberName)

	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "SELECT " & strDbTable & "Author.Author_ID From " & strDbTable & "Author WHERE " & strDbTable & "Author.Username='" & strMemberName & "';"
	
	'Query the database
	rsCommon.Open strSQL, adoCon
	
	'See if a user with that name is returned by the database
	If NOT rsCommon.EOF Then
		
		'Read in the user ID
		lngMemberID = CLng(rsCommon("Author_ID"))
		
		'Reset Server Objects
		rsCommon.Close
		Call closeDatabase()
	
		'Redirct to next page
		Response.Redirect("admin_user_permissions.asp?UID=" & lngMemberID & strQsSID3)
	
	'Else there is no user with that name returned so set an error code
	Else
	
		blnMemNotFound = true	
		
	End If


End If



'Reset Server Objects
Call closeDatabase()
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title>Create Member Permissions</title>
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
<!-- Check the from is filled in correctly before submitting -->
<script  language="JavaScript" type="text/javascript">
<!-- Hide from older browsers...

//Function to check form is filled in correctly before submitting
function CheckForm () {

	//Check for a group
	if (document.frmMessageForm.member.value==""){
		alert("Please enter a members username");
		return false;
	}
	
	return true
}


//Function to open pop up window
function winOpener(theURL, winName, scrollbars, resizable, width, height) {
	
	winFeatures = 'left=' + (screen.availWidth-10-width)/2 + ',top=' + (screen.availHeight-30-height)/2 + ',scrollbars=' + scrollbars + ',resizable=' + resizable + ',width=' + width + ',height=' + height + ',toolbar=0,location=0,status=1,menubar=0'
  	window.open(theURL, winName, winFeatures);
}
// -->
</script>
<!-- #include file="includes/admin_header_inc.asp" -->
<div align="center" class="text"><h1>Member Permissions</h1><br />
  <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
  <br />
  From here you can override  Group Permissions for different Members allowing  Members to have different permissions on forums, you can 
  also select Members to be able to have moderator privileges on forums.
  <p class="text">Select the Forum Member that you would like to Create, Edit, or Remove Member Permissions for.</p>
  <p></p>
</div>
<form action="admin_find_user.asp<% = strQsSID1 %>" method="post" name="frmMessageForm" id="frmMessageForm" onsubmit="return CheckForm();">
  <br />
  <table align="center" cellpadding="4" cellspacing="1" class="tableBorder" style="width:380px">
    <tr class="tableLedger">
      <td>Select a Member</td>
    </tr>
    <tr>
      <td class="tableRow">Username
        <input name="member" type="text" id="member" size="20" maxlength="25" value="<% = strMemberName %>" />
        <input type="submit" name="Submit" value="Next &gt;&gt;" />
        <a href="javascript:winOpener('pop_up_member_search.asp<% = strQsSID1 %>','memSearch',0,1,580,355)"><img src="<% = strImagePath %>member_search.png" title="Member Search" border="0" align="absbottom" /></a>
        <input type="hidden" name="postBack" value="true" />
      </td>
    </tr>
  </table>
</form>
<%

'If the username is already gone display an error message pop-up
If blnMemNotFound  Then
        Response.Write("<script  language=""JavaScript"">")
        Response.Write("alert('The Username entered could not be found.\n\nPlease check your spelling and try again.');")
        Response.Write("</script>")

End If 

%>
<!-- #include file="includes/admin_footer_inc.asp" -->
