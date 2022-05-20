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



'Dimension variables
Dim intCatID		'Holds the ID number of the category
Dim intSubID		'Holds the sub forum ID
Dim intMainForumID	'Holds the ID of the main forum is sub forum mode
Dim strForumName	'Hold sthe forum name




'Initilise variables
intCatID = 0
intForumID = 0



'Read in the details
intForumID = IntC(Request.QueryString("FID"))


'If this is a post back update the database
If Request.Form("postBack") AND  blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))

	'Read in the sub ID
	intMainForumID = IntC(Request.Form("mainForum"))

	'Get the Cat ID for this sub forum from the database
	strSQL = "SELECT " & strDbTable & "Forum.Cat_ID " & _
	"From " & strDbTable & "Forum " & _
	"WHERE " & strDbTable & "Forum.Forum_ID=" & intMainForumID & ";"

	'Query the database
	rsCommon.Open strSQL, adoCon

	'If there is a record get the Cat ID
	If NOT rsCommon.EOF Then intCatID = CInt(rsCommon("Cat_ID"))

	'Reset rsCommon
	rsCommon.Close


	'Update db
	strSQL = "UPDATE " & strDbTable & "Forum " & _
	"SET " & strDbTable & "Forum.Sub_ID=" & intMainForumID & ", " & strDbTable & "Forum.Cat_ID=" & intCatID & " " & _
	"WHERE " & strDbTable & "Forum.Forum_ID=" & intForumID & ";"

	adoCon.Execute(strSQL)	

	'Release server varaibles
	Call closeDatabase()

	Response.Redirect("admin_view_forums.asp" & strQsSID1)
End If


'Get the forum name

'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT " & strDbTable & "Forum.Forum_name " & _
"From " & strDbTable & "Forum " & _
"WHERE " & strDbTable & "Forum.Forum_ID=" & intForumID & ";"

'Query the database
rsCommon.Open strSQL, adoCon

'read in forum name
If Not rsCommon.EOF Then strForumName = rsCommon("Forum_name")

rsCommon.Close


'See if this forum has any sub forums

'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT " & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Cat_ID " & _
"From " & strDbTable & "Forum " & _
"WHERE " & strDbTable & "Forum.Sub_ID=" & intForumID & ";"

'Query the database
rsCommon.Open strSQL, adoCon
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title>Forum Details</title>
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
<!-- Check the from is filled in correctly before submitting -->
<script  language="JavaScript" type="text/javascript">

//Function to check form is filled in correctly before submitting
function CheckForm () {


	//Check for a forum
	if (document.frmNewForum.mainForum.value==""){
		alert("Please select a Main Forum for this Sub Forum is to be in");
		return false;
	}

	return true
}
</script>
<!-- #include file="includes/admin_header_inc.asp" -->
<div align="center"><h1>Change '<%= strForumName %>' Forum to Sub Forum </h1><br />
  <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
  <a href="admin_view_forums.asp<% = strQsSID1 %>">Return to the Forum Administration page</a><br />
  <br /><%

'If there is no main forum or cat then display a message
If NOT rsCommon.EOF Then
 	%>
  <table width="98%" border="0" cellspacing="0" cellpadding="1" height="135">
    <tr>
      <td align="center" class="text"><span class="lgText"> This Forum has Sub Forums, you must first remove any Sub Forums before '<%= strForumName %>' Forum can be changed into a Sub Forum.</span><br />
        <br />
        <a href="admin_view_forums.asp<% = strQsSID1 %>">Return to the Forum Administration page</a><br /></td>
    </tr>
  </table><%
 
 	'Close rs
 	rsCommon.Close
 
'Display page to chnage forum to sub forum
Else
	'Close rs
	rsCommon.Close
%>
</div>
<form action="admin_change_to_subforum.asp?FID=<% = intForumID %><% = strQsSID2 %>" method="post" name="frmNewForum" id="frmNewForum" onsubmit="return CheckForm();">
  <table align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
      <td colspan="2" class="tableLedger">Select Main Forum </td>
    </tr>
    <tr>
      <td colspan="2" class="tableRow">Select the Main Forum from the drop down list below that you would like '<%= strForumName %>' Forum to be a Sub Forum of.<br />
        <select name="mainForum">
          <option value=""<% If intSubID = 0 Then Response.Write(" selected") %>>-- Select Main Forum --</option><%
            
	'Read in the main forum name from the database
	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "SELECT " & strDbTable & "Category.Cat_ID, " & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Forum_name FROM " & strDbTable & "Category, " & strDbTable & "Forum WHERE " & strDbTable & "Category.Cat_ID=" & strDbTable & "Forum.Cat_ID AND " & strDbTable & "Forum.Sub_ID=0 ORDER BY " & strDbTable & "Category.Cat_order ASC, " & strDbTable & "Forum.Forum_Order ASC;"
	
	
	'Query the database
	rsCommon.Open strSQL, adoCon
	
	'Loop through all the main forums in the database
	Do while NOT rsCommon.EOF
	
		'Read in the deatils for the main forum
		strMainForumName = rsCommon("Forum_name")
		intMainForumID = CInt(rsCommon("Forum_ID"))
		
		'Make sure we are not showing this forum
		If intMainForumID <> intForumID Then 
	
			'Display a link in the link list to the cat
			Response.Write (vbCrLf & "		<option value=""" & intMainForumID & """")
			If intMainForumID = intSubID Then Response.Write(" selected")
			Response.Write(">" & strMainForumName & "</option>")
		End If
	
	
		'Move to the next record in the recordset
		rsCommon.MoveNext
	Loop
	
	'Close Rs
	rsCommon.Close
%>
      </select></td>
    </tr>
    <tr>
      <td colspan="2" align="center" class="tableBottomRow">
      <input type="hidden" name="postBack" id="postBack" value="true" />
      <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
      <input type="submit" name="Submit" value="Submit Forum Details"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
  </table>
  <br />
</form>
<%

End If

'Reset Server Objects
Call closeDatabase()
%>
<!-- #include file="includes/admin_footer_inc.asp" -->
