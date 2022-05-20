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
Dim strMode		'holds the mode of the page, set to true if changes are to be made to the database
Dim strForumDateFormat	'Holds the date format
Dim strForumYearFormat	'Holds the year format
Dim intForumTimeFormat	'Holds the time format
Dim strDateSeporator	'Holds the date seporator between the day/month/year
Dim saryMonth(12)	'Array holding each of the months
Dim strMorningID	'Holds the identifier to show for morning in 12 hour clock
Dim strAfternoonID	'Holds the identifier to show for afternoon in 12 hour clock
Dim intMonthLoopCounter	'Loop counter for the months
Dim strForumTimeOffSet	'Time of set (+) or (-)
Dim intForumTimeOffSet	'Time off set number
Dim lngLoopCounter	'Loop counter





blnBoldToday = BoolC(Request.Form("todayBold"))


'If the user is changing tthe colours then update the database
If Request.Form("postBack") AND blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))

	Call addConfigurationItem("Date_today_bold", blnBoldToday)
	
	'Update variables
	Application.Lock
	Application(strAppPrefix & "blnBoldToday") = CBool(blnBoldToday)
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
	blnBoldToday = CBool(getConfigurationItem("Date_today_bold", "bool"))
End If


'Reset Server Objects
rsCommon.Close





'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT " & strDbTable & "DateTimeFormat.* " & _
"From " & strDbTable & "DateTimeFormat " & _
"WHERE " & strDbTable & "DateTimeFormat.ID = 1;" 'Where cluase put for myODBC bug work around

'Set the cursor type property of the record set to Forward Only
rsCommon.CursorType = 0

'Set the Lock Type for the records so that the record set is only locked when it is updated
rsCommon.LockType = 3

'Query the database
rsCommon.Open strSQL, adoCon

'If the user is changing the date/time setup then update the database
If Request.Form("postBack") Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))

	With rsCommon
		'Update the recordset
		.Fields("Date_Format") = Request.Form("dateFormat")
		.Fields("Year_format") = Request.Form("yearFormat")
		.Fields("Time_format") = Request.Form("timeFormat")
		If Request.Form("seporator") = "space" Then .Fields("Seporator") = "&nbsp;" Else .Fields("Seporator") =  Request.Form("seporator")
		.Fields("am") = Request.Form("am")
		.Fields("pm") = Request.Form("pm")
		.Fields("Time_offset") = Request.Form("serverOffSet")
		.Fields("Time_offset_hours") = IntC(Request.Form("serverOffSetHours"))

		'Upadet the months (arrays start at 0 in VBScript but for simplisity we are not using location 1)
		For intMonthLoopCounter = 1 to 12
			.Fields("Month" & intMonthLoopCounter) = Request.Form("month" & intMonthLoopCounter)
		Next

		'Update the database with the new user's details
		.Update

		'Re-run the query to read in the updated recordset from the database
		.Requery
	End With
	
	
	'Empty the application level array holding the date and time format so that any changes are visable in the main forum
	Application(strAppPrefix & "saryAppDateTimeFormatData") = null
End If

'Read in the deatils from the database
If NOT rsCommon.EOF Then

	'Read in the date/time setup from the database
	'Update the recordset
	strForumDateFormat = rsCommon("Date_Format")
	strForumYearFormat = rsCommon("Year_format")
	intForumTimeFormat = CInt(rsCommon("Time_format"))
	strDateSeporator = rsCommon("Seporator")
	strMorningID = rsCommon("am")
	strAfternoonID = rsCommon("pm")
	strForumTimeOffSet = rsCommon("Time_offset")
	intForumTimeOffSet = CInt(rsCommon("Time_offset_hours"))

	'Update the months (arrays start at 0 in VBScript but for simplisity we are not using location 1)
	For intMonthLoopCounter = 1 to 12
		saryMonth(intMonthLoopCounter) = rsCommon.Fields("Month" & intMonthLoopCounter)
	Next
End If

'Include the time date function here so it's updated after any database update
%><!--#include file="functions/functions_date_time_format.asp" --><%

'Reset Server Objects
rsCommon.Close
Call closeDatabase()
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title>Date and Time Settings</title>

<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write("<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Forums(TM) ver. " & strVersion & "" & _
vbCrLf & "Info: http://www.webwizforums.com" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->" & vbCrLf)
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>
<!-- Check the from is filled in correctly before submitting -->
<script  language="JavaScript" type="text/javascript">
function CheckForm () {

	//Intialise variables
	var errorMsg = "";
	var errorMsgLong = "";

	//Check for all the month fields having values
	for (var count = 8; count <= 20; ++count){
		if (document.frmDateTime.elements[count].value == ""){

			var monthName;

			//get the month
			if (count == 8) {monthName = "January\t";}
			else if (count == 9) {monthName = "February\t";}
			else if (count == 10) {monthName = "March\t";}
			else if (count == 11) {monthName = "April\t";}
			else if (count == 12) {monthName = "May\t";}
			else if (count == 13) {monthName = "June\t";}
			else if (count == 14) {monthName = "July\t";}
			else if (count == 15) {monthName = "August\t";}
			else if (count == 16) {monthName = "September";}
			else if (count == 17) {monthName = "October\t";}
			else if (count == 18) {monthName = "Nevember";}
			else if (count == 19) {monthName = "December";}

			//Wriet the error message
			errorMsg += "\n" + monthName + " \t- Enter a value for " + monthName;
		}
	}

	//If there is aproblem with the form then display an error
	if ((errorMsg != "") || (errorMsgLong != "")){
		msg = "___________________________________________________________________\n\n";
		msg += "Your settings have not been updated because there are problem(s) with the form.\n";
		msg += "Please correct the problem(s) and re-submit the form.\n";
		msg += "___________________________________________________________________\n\n";
		msg += "The following field(s) need to be corrected: -\n";

		errorMsg += alert(msg + errorMsg + "\n" + errorMsgLong);
		return false;
	}

	return true;
}
</script>
<!-- #include file="includes/admin_header_inc.asp" -->
<div align="center">
  <h1>Forum Date and Time Settings</h1><br />
    <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
    <br />
    From here you can alter the forums global date and time format and alter the time off-set for a different time zone. <br />
    Registered members can have their own date a time settings by editing their individual Forum Profile.
</div>
<br />
<form action="admin_date_time_configure.asp<% = strQsSID1 %>" method="post" name="frmDateTime" id="frmDateTime" onsubmit="return CheckForm();">
  <table border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr align="left">
      <td height="30" colspan="2" class="tableLedger">Change Server Time </td>
    </tr>
    <tr>
      <td class="tableRow">Time offset from the web servers internal clock <br />
        <span class="smText"><br />
        Use this tool if you want the time shown in the forum to be different from the web servers current time. </span><br />
        <br />
        <span class="smText">eg. If it is 12.30pm on the web server and you want the time to show as 3.30pm you would select +3 hours. </span></td>
      <td class="tableRow"><select name="serverOffSet" id="serverOffSet">
          <option value="+"<% If strForumTimeOffSet = "+" Then Response.Write(" selected") %>>+</option>
          <option value="-"<% If strForumTimeOffSet = "-" Then Response.Write(" selected") %>>-</option>
        </select>
        <select name="serverOffSetHours">
          <%

	'Create list of time off-set
	For lngLoopCounter = 0 to 24
		Response.Write(VbCrLf & "      <option value=""" & lngLoopCounter & """")
		If intForumTimeOffSet = lngLoopCounter Then Response.Write("selected") 
		Response.Write(">" & lngLoopCounter & "</option>")
	Next
        
%>
        </select>
        Hours <br />
        <br />
        <span class="smText">Present web server date and time is:-<br />
        </span>
        <% = internationalDateTime(Now()) %>
        <br />
        <br />
        <span class="smText">What your settings have changed the time and date to:-<br />
        </span>
        <% 
     	Response.Write(stdDateFormat(Now(), True) & " at " & TimeFormat(Now()))
     	
     	%></td>
    </tr>
    <tr>
      <td colspan="2" class="tableLedger">Time Format </td>
    </tr>
    <tr>
      <td class="tableRow">Date Format:<br />
        <span class="smText">For 12 hour clock set to 12; for military time (24 hour clock) set to 24</span></td>
      <td class="tableRow"><select name="timeFormat">
          <option value="12" <% If intForumTimeFormat = 12 Then Response.Write("selected") %>>12 Hour Clock</option>
          <option value="24" <% If intForumTimeFormat = 24 Then Response.Write("selected") %>>24 Hour Clock</option>
        </select>
      </td>
    </tr>
    <tr>
      <td class="tableRow">Morning Identifier for 12 hour clock times:<br />
        <span class="smText">example: am</span></td>
      <td height="13" class="tableRow"><input type="text" name="am" maxlength="5" value="<% = strMorningID %>" size="5" />
      </td>
    </tr>
    <tr>
      <td class="tableRow">Afternoon Identifier for 12 hour clock times:<br />
        <span class="smText">example: pm</span></td>
      <td height="13" class="tableRow"><input type="text" name="pm" maxlength="5" value="<% = strAfternoonID %>" size="5" />
      </td>
    </tr>
    <tr>
      <td colspan="2" class="tableLedger">Date Format </td>
    </tr>
    <tr>
      <td width="62%" class="tableRow">Date Format:</td>
      <td width="38%" class="tableRow"><select name="dateFormat">
          <option value="dd/mm/yy" <% If strForumDateFormat = "dd/mm/yy" Then Response.Write("selected") %>>Day/Month/Year</option>
          <option value="mm/dd/yy" <% If strForumDateFormat = "mm/dd/yy" Then Response.Write("selected") %>>Month/Day/Year</option>
          <option value="yy/mm/dd" <% If strForumDateFormat = "yy/mm/dd" Then Response.Write("selected") %>>Year/Month/Day</option>
          <option value="yy/dd/mm" <% If strForumDateFormat = "yy/dd/mm" Then Response.Write("selected") %>>Year/Day/Month</option>
        </select>
      </td>
    </tr>
    <tr>
      <td width="62%" class="tableRow">Separator:<br />
        <span class="smText">This is the separator between the date eg: 12/12/2006, 12-12-2006, etc.</span></td>
      <td width="38%" class="tableRow"><select name="seporator">
          <option value="space"<% If strDateSeporator = "&nbsp;" OR strDateSeporator = Chr(32)Then Response.Write(" selected") %>>&lt;space&gt;</option>
          <option value="/"<% If strDateSeporator = "/" Then Response.Write(" selected") %>>/</option>
          <option value="\"<% If strDateSeporator = "\" Then Response.Write(" selected") %>>\</option>
          <option value="-"<% If strDateSeporator = "-" Then Response.Write(" selected") %>>-</option>
          <option value="&nbsp;-&nbsp;"<% If strDateSeporator = "&nbsp;-&nbsp;" Then Response.Write(" selected") %>>&nbsp;-&nbsp;</option>
          <option value="."<% If strDateSeporator = "." Then Response.Write(" selected") %>>.</option>
        </select>
      </td>
    </tr>
    <tr>
      <td width="62%" class="tableRow">Year Format:<br />
        <span class="smText">This is whether you want the date in 4 digits (2006) or in 2 digits (06)</span></td>
      <td width="38%" class="tableRow"><select name="yearFormat">
          <option value="long" <% If strForumYearFormat = "long" Then Response.Write("selected") %>>yyyy</option>
          <option value="short" <% If strForumYearFormat = "short" Then Response.Write("selected") %>>yy</option>
        </select>
      </td>
    </tr>
    <tr>
     <td class="tableRow">Display Dates From The Last 24 Hours in Bold:<br />
      <span class="smText">This allows any dates that are less than 24 hours old to be displayed in Bold.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="todayBold" value="True" <% If blnBoldToday = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="todayBold" value="False" <% If blnBoldToday = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
      <td width="62%"  height="2" class="tableRow">January*:<br />
        <span class="smText">This is what you would like displayed for January eg: 01, 1, Jan, etc. </span></td>
      <td width="38%" height="2" class="tableRow"><input type="text" name="month1" maxlength="15" value="<% = saryMonth(1) %>" size="15" />
      </td>
    </tr>
    <tr>
      <td width="62%" class="tableRow">February*:<br />
        <span class="smText">This is what you would like displayed for February eg: 02, 2, Feb, etc. </span></td>
      <td width="38%" class="tableRow"><input type="text" name="month2" maxlength="15" value="<% = saryMonth(2) %>" size="15" />
      </td>
    </tr>
    <tr>
      <td width="62%" class="tableRow">March*:</td>
      <td width="38%" class="tableRow"><input type="text" name="month3" maxlength="15" value="<% = saryMonth(3) %>" size="15" />
      </td>
    </tr>
    <tr>
      <td width="62%" class="tableRow">April*:</td>
      <td width="38%" class="tableRow"><input type="text" name="month4" maxlength="15" value="<% = saryMonth(4) %>" size="15" />
      </td>
    </tr>
    <tr>
      <td width="62%" class="tableRow">May*:</td>
      <td width="38%" class="tableRow"><input type="text" name="month5" maxlength="15" value="<% = saryMonth(5) %>" size="15" />
      </td>
    </tr>
    <tr>
      <td width="62%" class="tableRow">June*:</td>
      <td width="38%" class="tableRow"><input type="text" name="month6" maxlength="15" value="<% = saryMonth(6) %>" size="15" />
      </td>
    </tr>
    <tr>
      <td width="62%" class="tableRow">July*:</td>
      <td width="38%" class="tableRow"><input type="text" name="month7" maxlength="15" value="<% = saryMonth(7) %>" size="15" />
      </td>
    </tr>
    <tr>
      <td width="62%" class="tableRow">August*:</td>
      <td width="38%" class="tableRow"><input type="text" name="month8" maxlength="15" value="<% = saryMonth(8) %>" size="15" />
      </td>
    </tr>
    <tr>
      <td width="62%" class="tableRow">September*:</td>
      <td width="38%" class="tableRow"><input type="text" name="month9" maxlength="15" value="<% = saryMonth(9) %>" size="15" />
      </td>
    </tr>
    <tr>
      <td width="62%" class="tableRow">October*:</td>
      <td width="38%" class="tableRow"><input type="text" name="month10" maxlength="15" value="<% = saryMonth(10) %>" size="15" />
      </td>
    </tr>
    <tr>
      <td width="62%" class="tableRow">November*:</td>
      <td width="38%" class="tableRow"><input type="text" name="month11" maxlength="15" value="<% = saryMonth(11) %>" size="15" />
      </td>
    </tr>
    <tr>
      <td width="62%" class="tableRow">December*:</td>
      <td width="38%" class="tableRow"><input type="text" name="month12" maxlength="15" value="<% = saryMonth(12) %>" size="15" />
      </td>
    </tr>
    <tr align="center">
      <td height="2" colspan="2" class="tableBottomRow" >
          <input type="hidden" name="postBack" value="true" />
          <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
          <input type="submit" name="Submit" value="Update Date and Time Formats"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
          <input type="reset" name="Reset" value="Reset Form" />
       </td>
    </tr>
  </table>
</form>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
