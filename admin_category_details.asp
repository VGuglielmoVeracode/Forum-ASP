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
Dim strCatName		'Holds the name of the category
Dim intCatID		'Holds the ID number of the category
Dim intCatOrderNum	'Holds the category order number
 

'Read in the details
intCatID = IntC(Request.QueryString("CatID"))

'Count the number of categories
strSQL = "SELECT Count(" & strDbTable & "Category.Cat_ID) AS catCount " & _
"FROM " & strDbTable & "Category;"
		
'Query the database
rsCommon.Open strSQL, adoCon
		
'Get the number of forums in cat
If NOT rsCommon.EOF Then intCatOrderNum = CInt(rsCommon("catCount")) + 1
		
'Close recordset
rsCommon.Close
	


'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT " & strDbTable & "Category.* " & _
"From " & strDbTable & "Category " & _
"WHERE " & strDbTable & "Category.Cat_ID = " & intCatID & ";"

'Set the Lock Type for the records so that the record set is only locked when it is updated
rsCommon.LockType = 3

	
'Query the database
rsCommon.Open strSQL, adoCon


'If this is a post back then save the category
If BoolC(Request.Form("postBack")) AND blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
	

	'If this is a new one add new
	If intCatID = 0 Then rsCommon.AddNew

	'Update the recordset
	rsCommon.Fields("Cat_name") = Request.Form("category")
	If intCatID = 0 Then rsCommon.Fields("Cat_order") = intCatOrderNum
					
	'Update the database with the category details
	rsCommon.Update
	
		
	'Release server varaibles
	rsCommon.Close
	Call closeDatabase()
		
	Response.Redirect("admin_view_forums.asp" & strQsSID1)
		
	'Re-run the query to read in the updated recordset from the database
	rsCommon.Requery	
End If

'Read in the forum details from the recordset
If NOT rsCommon.EOF Then
	
	'Read in the forums from the recordset
	intCatID = CInt(rsCommon("Cat_ID"))
	strCatName = rsCommon("Cat_name")
End If

'Release server varaibles
rsCommon.Close
Call closeDatabase()
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title>Category Details</title>
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

	//Check for a a category
	if (document.frmNewForum.category.value==""){
		alert("Please enter the Category");
		return false;
	}
	
	return true
}
// -->
</script>
<!-- #include file="includes/admin_header_inc.asp" -->
<div align="center"><h1>Category</h1><br />
  <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
  <a href="admin_view_forums.asp<% = strQsSID1 %>">Return to the Category and Forum Set up and Admin page</a><br />
</div>
<form action="admin_category_details.asp?CatID=<% = intCatID %><% = strQsSID2 %>" method="post" name="frmNewForum" id="frmNewForum" onsubmit="return CheckForm();">
  <br />
  <table border="0" align="center" class="tableBorder" cellpadding="4" cellspacing="1" style="width: 380px">
    <tr>
      <td colspan="2" class="tableLedger">Category Details</td>
    </tr>
    <tr>
      <td width="33%" class="tableRow">Category*</td>
      <td width="67%" class="tableRow"><input type="text" name="category" maxlength="50" size="30" value="<% = Server.HTMLEncode(strCatName) %>" /></td>
    </tr>
    <tr>
      <td colspan="2" align="center" class="tableBottomRow">
      	<input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
      	<input type="hidden" name="postBack" value="true" />
        <input name="Submit" type="submit" id="Submit" value="Submit Category"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
  </table>
  <br />
</form>
<!-- #include file="includes/admin_footer_inc.asp" -->
