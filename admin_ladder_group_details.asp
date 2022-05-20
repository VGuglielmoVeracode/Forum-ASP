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
Dim intLadderGroup	'Holds the group ID
Dim strLadderGroupName	'Holds the name of the group
Dim lngMinimumPoints	'Holds the minimum amount of posts to be in that group
Dim blnSpecialGroup	'Set to true if a special group
Dim intStars		'Holds the number of stars for the group
Dim strCustomStars	'Holds the custom stars image if there is one fo0r this group
Dim intCatID		'Holds the cat ID
Dim sarryForums
Dim intCurrentRecord
Dim sarrySubForums
Dim intCurrentRecord2
Dim intSubForumID
Dim blnLadderGroup





'Read in the details
intLadderGroup = IntC(Request.QueryString("LID"))





'If this is a post back update the database
If Request.Form("postBack") AND blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
	
	'Read the various LadderGroup from the database
	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "SELECT " & strDbTable & "LadderGroup.* " & _
	"FROM " & strDbTable & "LadderGroup " & _
	"WHERE " & strDbTable & "LadderGroup.Ladder_ID = " & intLadderGroup & ";"
	
	
	'Set the cursor type property of the record set to Forward Only
	rsCommon.CursorType = 0
	
	'Set the Lock Type for the records so that the record set is only locked when it is updated
	rsCommon.LockType = 3
	
	'Query the database
	rsCommon.Open strSQL, adoCon

	
	'Read in the group details
	strLadderGroupName = Request.Form("LadderName")


	With rsCommon
		'If this is a new one add new
		If intLadderGroup = 0 Then .AddNew

		'Update the recordset
		.Fields("Ladder_Name") = strLadderGroupName

		'Update the database with the group details
		.Update
	End With
	

	'Close RS
	rsCommon.Close
	
	'Release server varaibles
	Call closeDatabase()

	'Redirect now update is complate
	Response.Redirect("admin_view_ladder_groups.asp" & strQsSID1)

End If 




'If this is an edit read in te Ladder group details
If intLadderGroup > 0 Then 
	
	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "" & _
	"SELECT " & strDbTable & "LadderGroup.* " & _
	"FROM " & strDbTable & "LadderGroup " & _
	"WHERE " & strDbTable & "LadderGroup.Ladder_ID = " & intLadderGroup & ";"
	
	'Query the database
	rsCommon.Open strSQL, adoCon

	If NOT rsCommon.EOF Then

		'Get the Ladder database
		strLadderGroupName = rsCommon("Ladder_Name")
	End If
	
	'Close the rs
	rsCommon.Close
	
End If



%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Ladder Group Details</title>
<meta name="generator" content="Web Wiz Forums" />
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

<script  language="JavaScript" type="text/javascript">

//Function to check form is filled in correctly before submitting
function CheckForm () {

	//Check for a group name
	if (document.frmGroup.GroupName.value==""){
		alert("Please select the Ladder Name for this Ladder Group");
		return false;
	}

	return true
}
</script>
<!-- #include file="includes/admin_header_inc.asp" -->
<div align="center"><h1>Ladder Group</h1><br />
  <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
  <a href="admin_view_ladder_groups.asp<% = strQsSID1 %>">Return to the Ladder Group Administration page</a><br />
</div>
<form action="admin_ladder_group_details.asp?LID=<% = intLadderGroup %><% = strQsSID2 %>" method="post" name="frmGroup" id="frmGroup" onsubmit="return CheckForm();">
  <br />
  <table border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder" style="width: 500px;">
    <tr>
      <td colspan="2" class="tableLedger">Ladder Group Details</td>
    </tr>
    <tr>
      <td width="50%" class="tableRow">Ladder Name*:</td>
      <td width="50%" valign="top" class="tableRow"><input name="LadderName" type="text" id="LadderName" value="<% = strLadderGroupName %>" size="25" maxlength="25"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      </td>
    </tr>
   </table>
   <br />
   <table width="100%" border="0" cellspacing="0" cellpadding="4">
    <tr>
      <td align="center"><input name="postBack" type="hidden" id="postBack" value="true" />
      	<input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
        <input type="submit" name="Submit" value="Update Ladder Group"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      </td>
    </tr>
  </table>
</form>
   <br /><%  


'Reset Server Objects
Call closeDatabase()

%>
<!-- #include file="includes/admin_footer_inc.asp" -->
