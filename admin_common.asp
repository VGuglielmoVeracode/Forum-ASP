<!-- #include file="includes/global_variables_inc.asp" -->
<!-- #include file="includes/setup_options_inc.asp" -->
<!-- #include file="includes/version_inc.asp" -->
<!-- #include file="database/database_connection.asp" -->
<!-- #include file="functions/functions_login.asp" -->
<!-- #include file="functions/functions_filters.asp" -->
<!-- #include file="functions/functions_common.asp" -->
<!-- #include file="functions/functions_session_data.asp" -->
<!-- #include file="functions/functions_hash1way.asp" -->
<!-- #include file="language_files/language_file_inc.asp" -->
<!-- #include file="functions/functions_windows_authentication.asp" -->
<!-- #include file="functions/functions_member_API.asp" -->
<!-- #include file="functions/functions_report_errors.asp" -->
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





'Set the timeout of the forum
Server.ScriptTimeout = 90

'Set the date time format to your own if you are getting a CDATE error
'Session.LCID = 1033



'Intialise variables
Dim strAdminReferer
Const strLSalt = "5CB237B1D85"
Const strCodeField = "&#076;_&#099;&#111;&#100;&#101;"
Const strCodeField2 = "&#065;_&#099;&#111;&#100;&#101;"
Const blnMassMailier = True
blnSslEnabledPage = True





'******************************************
'***  	   Database connection         ****
'******************************************

Call openDatabase(strCon)



'******************************************
'***    Read in Configuration Data     ****
'******************************************

Call getForumConfigurationData()


'******************************************
'***  		 Get Session ID        ****
'******************************************

'Call sub to get session data if not a searh engine spider
If NOT strOSType = "Search Robot" Then Call getSessionData() 
	


'******************************************
'***    Read in Logged-in User Data    ****
'******************************************

'Call the sub procedure to read in the details for this user
Call getUserData("AID")



'Check the refer is the same as the website the user logged in on
'Get the refer
strAdminReferer = LCase(Request.ServerVariables("HTTP_REFERER"))
		
'Trim the referer down to size
strAdminReferer = Replace(strAdminReferer, "http://", "")
strAdminReferer = Replace(strAdminReferer, "https://", "")
If NOT strAdminReferer = "" Then strAdminReferer = Mid(strAdminReferer, 1, InStr(strAdminReferer, "/")-1)
If Len(strAdminReferer) > 25 Then strAdminReferer = Mid(strAdminReferer, 1 ,25)
		
'Save the refer into teh session
If NOT getSessionItem("REF") = strAdminReferer Then
	
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("insufficient_permission.asp?M=sID" & strQsSID3)
End If



'If the user is not the admin or not logged in send them away
If intGroupID <> 1 Then 
	
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If




%>