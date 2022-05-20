<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="admin_common.asp" -->
<!--#include file="functions/functions_format_post.asp" -->
<!--#include file="includes/emoticons_inc.asp" -->
<!--#include file="functions/functions_date_time_format.asp" -->
<!--#include file="functions/functions_edit_post.asp" -->
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

'Dimension variables
Dim strEmail                    'Holds the new users e-mail address
Dim intUsersGroupID             'Holds the users group ID
Dim blnShowEmail                'Boolean set to true if the user wishes there e-mail address to be shown
Dim strHomepage                 'Holds the new users homepage if they have one
Dim strAvatar                   'Holds the avatar image
Dim strCheckUsername            'Holds the usernames from the database recordset to check against the new users requested username
Dim blnAutoLogin                'Boolean set to true if the user wants auto login trured on
Dim strImageFileExtension       'holds the file extension
Dim blnAccountReactivate        'Set to true if the users account needs to be reactivated
Dim blnSentEmail                'Set to true if the e-mail has been sent
Dim strEmailBody                'Holds the body of the welcome message e-mail
Dim strSubject                  'Holds the subject of the e-mail
Dim strSignature                'Holds the signature
Dim strICQNum                   'Holds the users ICQ Number
Dim strAIMAddress               'Holds the users AIM address
Dim strMSNAddress               'Holds the users MSN address
Dim strYahooAddress             'Holds the users Yahoo Address
Dim strOccupation               'Holds the users Occupation
Dim strInterests                'Holds the users Interests
Dim dtmDateOfBirth              'Holds the users Date Of Birth
Dim blnPMNotify                 'Set to true if the user want email notification of PM's
Dim strSmutWord                 'Holds the smut word to give better performance so we don't need to keep grabbing it form the recordset
Dim strSmutWordReplace          'Holds the smut word to be replaced with
Dim strMode                     'Holds the mode of the page
Dim blnEmailOK                  'Set to true if e-mail is not already in the database
Dim blnUsernameOK               'Set to true if the username requested does not already exsist
Dim intForumStartingGroup       'Holds the forum starting group ID number
Dim strEncryptedPassword         'Holds the encrypted password
Dim blnPasswordChange           'Holds if the password is changed or not
Dim blnEmailBlocked             'set to true if the email address is blocked
Dim strCheckEmailAddress        'Holds the email address to be checked
Dim lngUserProfileID            'Holds the users ID of the profile to get
Dim blnAdminMode                'Set to true if admin mode is enabled to update other members profiles
Dim blnUserActive               'Set to true if the users membership is active
Dim lngPosts                    'Holds the number of posts the user has made
Dim intDOBYear			'Holds the year of birth
Dim intDOBMonth			'Holds the month of birth
Dim intDOBDay			'Holds the day of birth
Dim strRealName			'Holds the persons real name
Dim strMemberTitle		'Holds the members title
Dim dtmServerTime		'Holds the current server time
Dim lngLoopCounter		'Holds the generic loop counter for page
Dim intUpdatePartNumber		'If an update holds which part to update
Dim strUsername			'Holds the users username
Dim strPassword			'Holds the usres password
Dim strUserCode			'Holds the users ID code
Dim strSalt			'Holds the salt value for the new member
Dim blnSuspended
Dim strAdminNotes
Dim blnNewsletter		'Set to true if newsletters are selected
Dim strGender			'Holds the users gender
Dim strTempUsername		'Holds atemp username for the user
Dim lngMemberPoints		'Holds the number of points the user has

'Initlise variables
blnUsernameOK = True
lngPosts = 0
lngMemberPoints = 0
blnUserActive = True
blnNewsletter = False
lngMemberPoints = 0
strDateFormat = saryDateTimeData(1,0)




'See if we are editing a user
lngUserProfileID = LngC(Request("PF"))


'If we have a ID number then put in adit mode
If lngUserProfileID <> 0 Then 
	strMode = "edit"
Else
	strMode = "new"
End If

'If the Profile has already been edited then update the Profile
If Request.Form("postBack") AND blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))


	'******************************************
	'***  Read in member details from form	***
	'******************************************

        'Read in the users details from the form
        strUsername = Trim(Mid(Request.Form("name1"), 1, 20))
        strPassword = Trim(Mid(Request.Form("password1"), 1, 20))
	strEmail = Trim(Mid(Request.Form("email"), 1, 60))
	strRealName = Trim(Mid(Request.Form("realName"), 1, 27))
	strGender = Trim(Mid(Request.Form("gender"), 1, 10))
	strHomepage = Trim(Mid(Request.Form("homepage"), 1, 48))
	strSignature = Mid(Request.Form("signature"), 1, 200)
	blnAttachSignature = BoolC(Request.Form("attachSig")) 
	'Check that the ICQ number is a number before reading it in
	If isNumeric(Request.Form("ICQ")) Then strICQNum = Trim(Mid(Request.Form("ICQ"), 1, 15))
	blnShowEmail = BoolC(Request.Form("emailShow"))
	blnPMNotify = BoolC(Request.Form("pmNotify"))
	blnAutoLogin = BoolC(Request.Form("Login"))
	strDateFormat = Trim(Mid(Request.Form("dateFormat"), 1, 10))
	strTimeOffSet = Trim(Mid(Request.Form("serverOffSet"), 1, 1))
	intTimeOffSet = IntC(Request.Form("serverOffSetHours"))
	blnReplyNotify = BoolC(Request.Form("replyNotify"))
	blnWYSIWYGEditor = BoolC(Request.Form("ieEditor"))
	blnUserActive = BoolC(Request.Form("active"))
        intUsersGroupID = IntC(Request.Form("group"))
        lngPosts = LngC(Request.Form("posts"))
        lngMemberPoints = LngC(Request.Form("points"))
        strMemberTitle = Trim(Mid(Request.Form("memTitle"), 1, 40))
        blnSuspended = BoolC(Request.Form("banned"))
        strAdminNotes = Trim(Mid(removeAllTags(Request.Form("notes")), 1, 255))
        If blnWebWizNewsPad Then blnNewsletter = BoolC(Request.Form("newsletter"))



        '******************************************
	'***     Read in the avatar		***
	'******************************************

       strAvatar = Trim(Mid(Request.Form("txtAvatar"), 1, 95))

       'If the avatar text box is empty then read in the avatar from the list box
       If strAvatar = "http://" OR strAvatar = "" Then strAvatar = Trim(Request.Form("SelectAvatar"))

       'If there is no new avatar selected then get the old one if there is one
       If strAvatar = "" Then strAvatar = Request.Form("oldAvatar")

       'If the avatar is the blank image then the user doesn't want one
       If strAvatar = strImagePath & "blank.gif" Then strAvatar = ""
        


        '******************************************
	'***     Clean up member details	***
	'******************************************

        'Clean up user input
        
        
	strRealName = removeAllTags(strRealName)
	strRealName = formatInput(strRealName)
	strGender = removeAllTags(strGender)
	strGender = formatInput(strGender)
	        
	'Call the function to format the signature
	strSignature = FormatPost(strSignature)
	
	'Call the function to format forum codes
	strSignature = FormatForumCodes(strSignature)
	
	 'Call the filters to remove malcious HTML code
	 strSignature = HTMLsafe(strSignature)
	 
	 'Trim signature down to a 255 max characters to prevent database errors
	 strSignature = Mid(strSignature, 1, 255)
	
	
	'If the user has not entered a hoempage then make sure the homepage variable is blank
	If strHomepage = "http://" Then strHomepage = ""
	
            
	strMemberTitle = removeAllTags(strMemberTitle) 
	strMemberTitle = formatInput(strMemberTitle)
	
	

	'******************************************
	'***     Check the avatar is OK		***
	'******************************************
	'If there is no . in the link then there is no extenison and so can't be an image
        If inStr(1, strAvatar, ".", 1) = 0 Then
                  strAvatar = ""
               
         'Else remove malicious code and check the extension is an image extension
         Else
                'Call the filter for the image
                strAvatar = formatInput(strAvatar)
         End If
        

	'******************************************
	'***     Check the username is OK	***
	'******************************************


        'Check there is a username
      	If strUsername = "" Then blnUsernameOK = False


        '******************************************
	'***   Check the username is availabe	***
	'******************************************

        'If the username is not already written off then check it's not already gone
        If blnUsernameOK Then
        	
        	'Make username SQL safe
       		strTempUsername = formatSQLInput(strUsername)

         	'Read in the the usernames from the database to check that the username does not already exsist

                'Initalise the strSQL variable with an SQL statement to query the database
                 strSQL = "SELECT " & strDbTable & "Author.Username " & _
                 "FROM " & strDbTable & "Author" & strDBNoLock & "  " & _
                 "WHERE " & strDbTable & "Author.Username = '" & strTempUsername & "' AND " & strDbTable & "Author.Author_ID <> " & lngUserProfileID & ";"

                'Set the cursor type property of the record set to Forward Only
	        rsCommon.CursorType = 0
	
	        'Set the Lock Type for the records so that the record set is only locked when it is updated
	        rsCommon.LockType = 3
	
	        'Open the author table
	        rsCommon.Open strSQL, adoCon

                'If there is a record returned from the database then the username is already used
                If NOT rsCommon.EOF Then 
                	blnUsernameOK = False
                End If

		'Close rs
                rsCommon.Close
 	End If
		

	'******************************************
	'*** 	     Create a usercode 		***
	'******************************************

        'Calculate a code for the user
        strUserCode = userCode(strUsername)

	

	'******************************************
	'*** 		Encrypt password	***
	'******************************************

        'Encrypt password
	If strPassword <> "" Then
		
		'Encrypt password
		If blnEncryptedPasswords Then																							
	
			'Genrate a slat value
		       	strSalt = getSalt(Len(strPassword))
		
		       'Concatenate salt value to the password
		       strEncryptedPassword = strPassword & strSalt
		
		       'Encrypt the password
		       strEncryptedPassword = HashEncode(strEncryptedPassword)
		
		'Else the password is not set to be encrypted so place the un-encrypted password into the strEncryptedPassword variable
		Else
	
			strEncryptedPassword = strPassword
		End If
	 End If



	'Intialise the strSQL variable with an SQL string to open a record set for the Author table
        strSQL = "SELECT " & strDbTable & "Author.Author_ID, " & strDbTable & "Author.Group_ID, " & strDbTable & "Author.Username, " & strDbTable & "Author.Real_name, " & strDbTable & "Author.Gender, " & strDbTable & "Author.User_code, " & strDbTable & "Author.Password, " & strDbTable & "Author.Salt, " & strDbTable & "Author.Author_email, " & strDbTable & "Author.Homepage, " & strDbTable & "Author.Location, " & strDbTable & "Author.MSN, " & strDbTable & "Author.Yahoo, " & strDbTable & "Author.ICQ, " & strDbTable & "Author.AIM, " & strDbTable & "Author.Occupation, " & strDbTable & "Author.Interests, " & strDbTable & "Author.DOB, " & strDbTable & "Author.Signature, " & strDbTable & "Author.No_of_posts, " & strDbTable & "Author.Points, " & strDbTable & "Author.No_of_PM, " & strDbTable & "Author.Join_date, " & strDbTable & "Author.Avatar, " & strDbTable & "Author.Avatar_title, " & strDbTable & "Author.Last_visit, " & strDbTable & "Author.Time_offset, " & strDbTable & "Author.Time_offset_hours, " & strDbTable & "Author.Date_format, " & strDbTable & "Author.Show_email, " & strDbTable & "Author.Attach_signature, " & strDbTable & "Author.Active, " & strDbTable & "Author.Rich_editor, " & strDbTable & "Author.Reply_notify, " & strDbTable & "Author.PM_notify, " & strDbTable & "Author.Skype, " & strDbTable & "Author.Login_attempt, " & strDbTable & "Author.Banned, " & strDbTable & "Author.Info, " & strDbTable & "Author.Newsletter " &_
	"FROM " & strDbTable & "Author" & strRowLock & " " & _
	"WHERE " & strDbTable & "Author.Author_ID = " & lngUserProfileID & ";"

        'Set the cursor type property of the record set to Forward Only
        rsCommon.CursorType = 0

        'Set the Lock Type for the records so that the record set is only locked when it is updated
        rsCommon.LockType = 3

        'Open the author table
        rsCommon.Open strSQL, adoCon


	'******************************************
	'*** 	  	Update datbase		***
	'******************************************

        'If this is new reg and the username and email is OK or this is an update then register the new user or update the rs
        If blnUsernameOK Then

            
                'Insert the user's details into the rs
                With rsCommon
                
                	'If not an edit then addnew
                	If strMode <> "edit" Then .AddNew
                        
                        'Don't update username is member API is enabled  
                        If blnMemberAPI = False Then .Fields("Username") = strUsername
                        
                        

                        
                        'If new undate the fields below
                        If strMode = "new" Then
	                        .Fields("Join_date") = internationalDateTime(Now())
				.Fields("Last_visit") = internationalDateTime(Now())
			End If
			
			
			'If the password has changed update it
			If strPassword <> "" AND lngUserProfileID <> 2 Then 
	                        .Fields("Password") = strEncryptedPassword
		                .Fields("Salt") = strSalt
		        End If
	                
	                
	                'If editing and windows authentication is enabled don't update the User_code
	                If strMode = "edit" Then
		                If blnWindowsAuthentication = False AND lngUserProfileID > 2 Then .Fields("User_code") = strUserCode
		       
		        'Else this is a new user so create a User_code for them
		        Else
		        	 .Fields("User_code") = strUserCode
		        End If
	                
	                	
	                If lngUserProfileID <> 2 Then
	                	.Fields("Author_email") = strEmail
	        	End If
	        	
                        .Fields("Real_name") = strRealName
                        .Fields("Gender") = strGender
		       	.Fields("Avatar") = strAvatar
		        .Fields("Homepage") = strHomepage
		        .Fields("Signature") = strSignature
		        .Fields("Attach_signature") = blnAttachSignature
	             	.Fields("Date_format") = strDateFormat
			.Fields("Time_offset") = strTimeOffSet
 			.Fields("Time_offset_hours") = intTimeOffSet
	    		.Fields("Reply_notify") = blnReplyNotify
	          	.Fields("Rich_editor") = blnWYSIWYGEditor
	          	.Fields("PM_notify") = blnPMNotify
	       		.Fields("Show_email") = blnShowEmail 
                        	
			
			If blnWebWizNewsPad Then .Fields("Newsletter") = blnNewsletter
			
			
			'Admin bits
			If lngUserProfileID <> 1 AND lngUserProfileID <> 2 Then
				.Fields("Group_ID") = intUsersGroupID
				.Fields("Active") = blnUserActive
				.Fields("Banned") = blnSuspended
			End If	
                        .Fields("Avatar_title") = strMemberTitle
			.Fields("No_of_posts") = lngPosts
			.Fields("Points") = lngMemberPoints
			.Fields("Info") = strAdminNotes
                	

                        'Update the database with the new user's details (needed for MS Access which can be slow updating)
                        .Update

                        'Re-run the query to read in the updated recordset from the database
                        .Requery
                End With


		'******************************************
		'*** 	 	 Clean up   		***
		'******************************************

                'Reset server Object
                rsCommon.Close
                Call closeDatabase()


		'******************************************
		'*** 	 Redirect to message page	***
		'******************************************

                'Redirect the welcome new user page
                Response.Redirect("admin_added_member.asp?M=" & strMode & strQsSID3)
        End If
        
        rsCommon.Close
End If





'******************************************
'***     Get the user details from db	***
'******************************************

'If this is a profile update get the users details to update
If  strMode = "edit" Then

        'Read the various forums from the database
        'Initalise the strSQL variable with an SQL statement to query the database
        strSQL = "SELECT " & strDbTable & "Author.* " & _
	"FROM " & strDbTable & "Author" & strDBNoLock & " " & _
	"WHERE " & strDbTable & "Author.Author_ID = " & lngUserProfileID

        'Query the database
        rsCommon.Open strSQL, adoCon

        'If there is no matching profile returned by the recordset then redirect the user to the main forum page
        If rsCommon.EOF Then

                'Reset server Object
                rsCommon.Close
                Call closeDatabase()

                Response.Redirect("default.asp" & strQsSID1)
        End If

        'Read in the new user's profile from the recordset
        strUsername = rsCommon("Username")
        strRealName = rsCommon("Real_name")
        strGender = rsCommon("Gender")
        If NOT isNull(rsCommon("Author_email")) Then strEmail = formatInput(rsCommon("Author_email"))
        If blnWebWizNewsPad Then blnNewsletter = CBool(rsCommon("Newsletter"))
        blnShowEmail = CBool(rsCommon("Show_email"))
        If NOT isNull(rsCommon("Homepage")) Then strHomepage = formatInput(rsCommon("Homepage"))
        strSignature = rsCommon("Signature")
        strAvatar = rsCommon("Avatar")
        strMemberTitle = rsCommon("Avatar_title")
        strDateFormat = rsCommon("Date_format")
        strTimeOffSet = rsCommon("Time_offset")
        intTimeOffSet = CInt(rsCommon("Time_offset_hours"))
        blnReplyNotify = CBool(rsCommon("Reply_notify"))
        blnAttachSignature = CBool(rsCommon("Attach_signature"))
        blnWYSIWYGEditor = CBool(rsCommon("Rich_editor"))
	blnPMNotify = CBool(rsCommon("PM_notify"))
        

       
	intUsersGroupID = CInt(rsCommon("Group_ID"))
	blnUserActive = CBool(rsCommon("Active"))
	lngPosts = CLng(rsCommon("No_of_posts"))
	If isNumeric(rsCommon("Points")) Then lngMemberPoints = CLng(rsCommon("Points")) Else lngMemberPoints = 0
	blnSuspended = CBool(rsCommon("Banned"))
	strAdminNotes = rsCommon("Info")
       


        'Reset Server Objects
        rsCommon.Close

	
   
End If

'Covert the signature back to forum codes
If strSignature <> "" Then  strSignature = EditPostConvertion(strSignature)



%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title><% If strMode = "edit" Then Response.Write("Update Member") Else Response.Write("Register New User") %></title>
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
<script language="JavaScript" type="text/javascript">

//Function to check form is filled in correctly before submitting
function CheckForm () {

        //Initialise variables
        var errorMsg = "";
        var errorMsgLong = "";

        //Check for a username
        if (document.frmRegister.name1.value.length < <% = intMinUsernameLength %>){
                errorMsg += "\nUsername \t- Your Username must be at least <% = intMinUsernameLength %> characters";
        }<%

'If not in edit mode
If strMode <> "edit" AND lngUserProfileID <> 2 Then
%>
	
        //Check for a password
        if (document.frmRegister.password1.value.length < <% = intMinPasswordLength %>){
                errorMsg += "\nPassword \t- Your Password must be at least <% = intMinPasswordLength %> characters";
        }<%
End If

'don't display if guest account editing
If lngUserProfileID <> 2 Then
%>

        //Check both passwords are the same
        if ((document.frmRegister.password1.value) != (document.frmRegister.password2.value)){
                errorMsg += "\nPassword Error\t- The passwords entered do not match";
                document.frmRegister.password1.value = ""
                document.frmRegister.password2.value = ""
        }<%
        
        
'If this is an update only check the password length if the user is enetring a new password
ElseIf strMode <> "edit" AND blnWindowsAuthentication = False AND lngUserProfileID <> 2 Then

	%>
        //Check for a password
        if ((formArea.password1.value.length < <% = intMinPasswordLength %>) && (formArea.password1.value.length > 0)){
                errorMsg += "\n<% = strTxtErrorPasswordChar %>";
        }<%
End If

'don't display if guest account editing
If lngUserProfileID <> 2 Then
%>

        //If an e-mail is entered check that the e-mail address is valid
        if (document.frmRegister.email.value.length >0 && (document.frmRegister.email.value.indexOf("@",0) == -1||document.frmRegister.email.value.indexOf(".",0) == -1)) {
                errorMsg +="\nEmail\t\t- Enter your valid email address";
          errorMsgLong += "\n- If you don't want to enter your email address then leave the email field blank"; 
        }

        //Check to make sure the user is not trying to show their email if they have not entered one
        if (document.frmRegister.email.value == "" && document.frmRegister.emailShow[0].checked == true){
                errorMsgLong += "\n- You can not show your email address if you haven\'t entered one!";
                document.frmRegister.emailShow[1].checked = true
                document.frmRegister.email.focus();
        }<%
End If

%>
	
        //Check that the signature is not above 200 chracters
        if (document.frmRegister.signature.value.length > 200){
                errorMsg += "\nSignature \t- Your signature has to many characters";
                errorMsgLong += "\n- You have " + document.frmRegister.signature.value.length + " characters in your signature, you must shorten it to below 200";
        }

        //If there is aproblem with the form then display an error
        if ((errorMsg != "") || (errorMsgLong != "")){
                msg = "_______________________________________________________________\n\n";
                msg += "The form has not been submitted because there are problem(s) with the form.\n";
                msg += "Please correct the problem(s) and re-submit the form.\n";
                msg += "_______________________________________________________________\n\n";
                msg += "The following field(s) need to be corrected: -\n";

                errorMsg += alert(msg + errorMsg + "\n" + errorMsgLong);
                return false;
        }

        //Reset the submition action
        document.frmRegister.action = "admin_register.asp<% = strQsSID1 %>"
        document.frmRegister.target = "_self";

        return true;
}

//Function to count the number of characters in the signature text box
function DescriptionCharCount() {
        document.frmRegister.countcharacters.value = document.frmRegister.signature.value.length;
}

//Function to open pop up window
function winOpener(theURL, winName, scrollbars, resizable, width, height) {
	
	winFeatures = 'left=' + (screen.availWidth-10-width)/2 + ',top=' + (screen.availHeight-30-height)/2 + ',scrollbars=' + scrollbars + ',resizable=' + resizable + ',width=' + width + ',height=' + height + ',toolbar=0,location=0,status=1,menubar=0'
  	window.open(theURL, winName, winFeatures);
}

//Function to open preview post window
function OpenPreviewWindow(targetPage, formName){
	
	now = new Date  
	
	//Open the window first 	
   	winOpener('','preview',1,1,680,400)
   		
   	//Now submit form to the new window
   	formName.action = targetPage + "?ID=" + now.getTime();	
	formName.target = "preview";
	formName.submit();
}

//Function to count characters in textarea
function characterCounter(charNoBox, textFeild) {
	document.getElementById(charNoBox).value = document.getElementById(textFeild).value.length;
}
</script>
<!-- #include file="includes/admin_header_inc.asp" -->
<div align="center"><h1><% If strMode = "edit" Then Response.Write("Update Member") Else Response.Write("Register New User") %></h1><br />
  <span class="text"><a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
  <br /><%
  
'If member API enabled then display a message to the user
If blnMemberAPI AND strMode = "new" Then 
	Response.Write("  This option is not available if the Member API is enabled.<br /><br />After logging in through your own login system, members will be added to the forum using the Member API when they enter the forum.")
Else
	Response.Write("  From here you can register a new member or update current members on the forum.")
End If
%>
<br /><br /></span>
</div><%
  
'If member API enabled do not show the form
If blnMemberAPI = False OR strMode = "edit" Then 

%>
<form method="post" name="frmRegister" action="admin_register.asp<% = strQsSID1 %>" onReset="return confirm('Are you sure you want to reset the form?');">
<table width="98%" height="14" border="0" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
  <tr class="tableLedger">
    <td colspan="2">Registration Details</td>
  </tr>
  <tr >
    <td width="50%" class="tableRow">Username*<br />
      <span class="smText">This is the name displayed when you use the forum</span></td>
    <td width="50%" class="tableRow"><% 
    	
    	'If API is enabled display the name but no edit
    	If strMode ="edit" AND blnMemberAPI Then 
    		
    		Response.Write("<input name=""name1"" type=""hidden"" value=""" &  strUsername & """ />" & strUsername)
    	Else	
    		%><input name="name1" type="text" size="15" maxlength="20" value="<% = strUsername %>"<% If strMode ="edit" AND blnDemoMode Then Response.Write(" disabled=""disabled""") %> /><%
    	
	End If	
    %></td>
  </tr><%

	'don't display if guest account editing
	If lngUserProfileID <> 2 Then
	
%>
  <tr>
    <td width="50%" class="tableRow">Password*</td>
    <td width="50%" valign="top" class="tableRow"><input name="password1" type="password" size="15" maxlength="15" autocomplete="off"<% If strMode = "edit" AND blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
  </tr>
  <tr >
    <td width="50%"  height="2" class="tableRow">Retype Password*</td>
    <td width="50%" height="2" valign="top" class="tableRow"><input name="password2" type="password" size="15" maxlength="15" autocomplete="off"<% If strMode = "edit" AND blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
  </tr>
  <tr>
    <td width="50%" class="tableRow">Email Address <br />
      <span class="smText">Not required, but useful if you wish to be notified when someone answers one of your posts or if you lose your password.</span><br />    </td>
    <td width="50%" valign="top" class="tableRow"><input type="text" name="email" size="30" maxlength="60" value="<% = strEmail %>" />
    &nbsp;</td>
  </tr><%
   
		'If Newsletter is enabled
		If blnWebWizNewsPad Then
        		%>
   <tr class="tableRow">
    <td width="50%"><% = strTxtNewsletterSubscription %><br /><span class="smText"><% = strTxtSignupToRecieveNewsletters & " " & strWebsiteName %></span></td>
    <td width="50%" valign="top"><% = strTxtYes %><input type="radio" name="newsletter" id="newsletter" value="True" <% If blnNewsletter = True Then Response.Write "checked" %> />&nbsp;&nbsp;<% = strTxtNo %><input type="radio" name="newsletter" id="newsletter" value="False" <% If blnNewsletter = False Then Response.Write "checked" %> /></td>
   </tr><%
   
		End If
	End If
%>
  <tr class="tableLedger">
    <td colspan="2">Profile Information (not required)</td>
  </tr>
  <tr>
    <td width="50%" class="tableRow">Real Name</td>
    <td width="50%" class="tableRow"><input name="realName" type="text" value="<% = strRealName %>" size="30" maxlength="27" /></td>
  </tr>
  <tr class="tableRow">
    <td width="50%">Gender</td>
    <td width="50%">
     <select name="gender" id="gender" tabindex="11">
      <option value=""<% If strGender = "" Or strGender = null Then Response.Write(" selected") %>>Private</option>
      <option value="<% = strTxtMale %>"<% If strGender = strTxtMale Then Response.Write(" selected") %>><% = strTxtMale %></option>
      <option value="<% = strTxtFemale %>"<% If strGender = strTxtFemale Then Response.Write(" selected") %>><% = strTxtFemale %></option>
     </select>
    </td>
  </tr>
  <tr>
    <td width="50%" class="tableRow">Homepage</td>
    <td width="50%" class="tableRow"><input name="homepage" type="text" value="<% If strHomepage = "" Then Response.Write "http://" Else Response.Write(strHomepage) %>" size="30" maxlength="48" /></td>
  </tr>
  <tr>
    <td height="2" valign="top" class="tableRow">Select Avatar <br />
      <span class="smText">This is the small icon shown next to your posts. Type the path to your own avatar in to the text box provided (must be 64 x 64  pixels).</span></td>
    <td height="2" valign="top" class="tableRow" ><table width="290" border="0" cellspacing="0" cellpadding="1">
        <tr>
          <td width="168">&nbsp;</td>
          <td width="122" align="center"><img src="<% = strImagePath %>blank.gif" name="avatar" width="64" height="64" id="avatar" />
            <input type="hidden" name="oldAvatar" value="<% = strAvatar %>" /></td>
        </tr>
        <tr>
          <td width="168"><input type="text" name="txtAvatar" id="txtAvatar" size="30" maxlength="95" value="<%

		'If the avatar is the persons own then display the link
		If InStr(1, strAvatar, "http://") > 0 Then
			Response.Write(strAvatar)
		Else
			Response.Write("http://")
		End If
        %>" onchange="oldAvatar.value=''" />
          </td>
          <td width="122"><input type="button" name="preview" value="Preview" onclick="avatar.src = txtAvatar.value" />
          </td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td height="2" valign="top" class="tableRow">Signature<br />
      <span class="smText">Enter a signature that you would like shown at the bottom of your Forum Posts&nbsp;(max 200 characters)</span><br />
      <br />
      <br />
    <a href="javascript:winOpener('BBcodes.asp<% = strQsSID1 %>','codes',1,1,610,500)" class="smLink">Forum Codes</a> can be used in your signature</td>
    <td height="2" valign="top" class="tableRow"><textarea name="signature" cols="30" rows="3" onkeydown="DescriptionCharCount();" onkeyup="DescriptionCharCount();"><% = strSignature %></textarea>
      <br />
      <input size="3" value="0" name="countcharacters" maxlength="3" />
      <input onclick="DescriptionCharCount();" type="button" value="Character Count" name="Count" />
    &nbsp;&nbsp;<span class="smText"><a href="javascript:OpenPreviewWindow('signature_preview.asp<% = strQsSID1 %>', document.frmRegister)" class="smLink">Signature Preview</a> </td>
  </tr>
  <tr>
    <td width="50%" class="tableRow">Always attach my signature to posts</td>
    <td width="50%" valign="top" class="tableRow">Yes<input type="radio" name="attachSig" value="true" <% If blnAttachSignature = True Then Response.Write "checked" %> />&nbsp;&nbsp;No<input type="radio" name="attachSig" value="false" <% If blnAttachSignature = False Then Response.Write "checked" %> /></td>
  </tr>
  <tr class="tableLedger">
    <td colspan="2">Forum Preferences</td>
  </tr>
  <tr>
    <td width="50%" class="tableRow">Show my Email Address <br />
    Hide your email address if you want it kept private from other users.</td>
    <td width="50%" valign="top" class="tableRow">Yes<input type="radio" name="emailShow" id="emailShow" value="True" <% If blnShowEmail = True Then Response.Write "checked" %> />&nbsp;&nbsp;No<input type="radio" name="emailShow" id="emailShow" value="False" <% If blnShowEmail = False Then Response.Write "checked" %> /></td>
  </tr>
  <tr>
    <td width="50%" class="tableRow">Notify me of replies to posts <br />
      <span class="smText">Sends an email when someone replies to a topic you have posted in. This can be changed whenever you post.</span></td>
    <td width="50%" valign="top" class="tableRow">Yes<input type="radio" name="replyNotify" id="replyNotify" value="True" <% If blnReplyNotify = True Then Response.Write "checked" %> />&nbsp;&nbsp;No<input type="radio" name="replyNotify" id="replyNotify" value="False" <% If blnReplyNotify = False Then Response.Write "checked" %> /></td>
  </tr>
  <tr>
    <td width="50%" class="tableRow">Notify me by email when I receive a Private Message</td>
    <td width="50%" valign="top" class="tableRow">Yes<input type="radio" name="pmNotify" id="pmNotify" value="True" <% If blnPMNotify = True Then Response.Write "checked" %> />&nbsp;&nbsp;No<input type="radio" name="pmNotify" id="pmNotify" value="False" <% If blnPMNotify = False Then Response.Write "checked" %> /></td>
  </tr>
  <tr>
    <td width="50%" class="tableRow">Enable the WYSIWYG post editor <br />
      <span class="smText">Only browsers that are detected as being Rich Text Enabled will have this feature available when posting.</span></td>
    <td width="50%" valign="top" class="tableRow">Yes<input type="radio" name="ieEditor" id="ieEditor" value="True" <% If blnWYSIWYGEditor = True Then Response.Write "checked" %> />&nbsp;&nbsp;No<input type="radio" name="ieEditor" id="ieEditor" value="False" <% If blnWYSIWYGEditor = False Then Response.Write "checked" %> /></td>
  </tr>
  <tr>
    <td width="50%" class="tableRow">Automatically log me in when I return to the Forum</td>
    <td width="50%" valign="top" class="tableRow">Yes<input type="radio" name="Login" id="Login" value="True" <% If blnAutoLogin = True Then Response.Write "checked" %> />&nbsp;&nbsp;No<input type="radio" name="Login" id="Login" value="False" <% If blnAutoLogin = False Then Response.Write "checked" %> /></td>
  </tr>
  <tr>
    <td width="50%" class="tableRow">Time offset from forum time <br />
      <span class="smText">Present server date and time is:<%


	'Get the current server time
	dtmServerTime = Now()

	'Make sure that the time and date format function isn't effected by the server time off set
	If strTimeOffSet = "-" Then
		dtmServerTime = DateAdd("h", + intTimeOffSet, dtmServerTime)
	ElseIf strTimeOffSet = "+" Then
		dtmServerTime = DateAdd("h", - intTimeOffSet, dtmServerTime)
	End If

	'Display the current server time
	Response.Write(stdDateFormat(dtmServerTime, True) & " " & strTxtAt & " " & TimeFormat(dtmServerTime))

%></span></td>
    <td valign="top">
     <select name="serverOffSet" id="serverOffSet">
      <option value="+" <% If strTimeOffSet = "+" Then Response.Write("selected") %>>+</option>
      <option value="-" <% If strTimeOffSet = "-" Then Response.Write("selected") %>>-</option>
     </select>
    <select name="serverOffSetHours"><%

	'Create list of time off-set
	For lngLoopCounter = 0 to 24
		Response.Write(VbCrLf & "      <option value=""" & lngLoopCounter & """")
		If intTimeOffSet = lngLoopCounter Then Response.Write("selected")
		Response.Write(">" & lngLoopCounter & "</option>")
	Next

%>
     </select>
    hours</td>
  </tr>
  <tr>
    <td width="50%" class="tableRow">Date Format</td>
    <td width="50%" valign="top" class="tableRow">
     <select name="dateFormat">
      <option value="dd/mm/yy" <% If strDateFormat = "dd/mm/yy" Then Response.Write("selected") %>>Day/Month/Year</option>
      <option value="mm/dd/yy" <% If strDateFormat = "mm/dd/yy" Then Response.Write("selected") %>>Month/Day/Year</option>
      <option value="yy/mm/dd" <% If strDateFormat = "yy/mm/dd" Then Response.Write("selected") %>>Year/Month/Day</option>
      <option value="yy/dd/mm" <% If strDateFormat = "yy/dd/mm" Then Response.Write("selected") %>>Year/Day/Month</option>
     </select>
    </td>
  </tr>
  <tr class="tableLedger">
    <td colspan="2">Admin and Moderator Functions</td>
  </tr><%
 	'Only allow update if not the built in admin or guest account
 	If lngUserProfileID <> 1 AND lngUserProfileID <> 2 Then
%>
  <tr>
    <td width="50%" class="tableRow">User is active</td>
    <td width="50%" valign="top" class="tableRow">Yes<input type="radio" name="active" id="active" value="True" <% If blnUserActive = True Then Response.Write "checked" %>>&nbsp;&nbsp;No<input type="radio" name="active" id="active" value="False" <% If blnUserActive = False Then Response.Write "checked" %> /></td>
  </tr>
  <tr class="tableRow">
    <td width="50%">Suspend User</td>
    <td width="50%">Yes<input type="radio" name="banned" id="banned" value="True" <% If blnSuspended = True Then Response.Write "checked" %>>&nbsp;&nbsp;No<input type="radio" name="banned" id="banned" value="False" <% If blnSuspended = False Then Response.Write "checked" %> /></td>
   </tr>
  <%
     			'Get the forum groups from the database so admin can change the members group
	
	                'Initlise SQL query
	                strSQL = "" & _
	                "SELECT " & strDbTable & "Group.*, " & strDbTable & "LadderGroup.* " & _
			"FROM " & strDbTable & "Group " & _
			"LEFT JOIN " & strDbTable & "LadderGroup ON " & strDbTable & "Group.Ladder_ID = " & strDbTable & "LadderGroup.Ladder_ID " & _
			"ORDER BY " & strDbTable & "LadderGroup.Ladder_Name ASC, " & strDbTable & "Group.Minimum_posts ASC, " & strDbTable & "Group.Group_ID ASC;"
	
	                'Query the database
	                rsCommon.Open strSQL, adoCon
	
	                'If there are groups then disply them
	                If NOT rsCommon.Eof Then
		

     %>
  <tr>
    <td width="50%" class="tableRow">Group</td>
    <td width="50%" valign="top" class="tableRow"><select name="group">
        <%

	                        'Loop round to display all the groups
	                        Do While NOT rsCommon.EOF
	
	                                Dim intSelGroupID
	                                Dim strSelGroupName
	                                Dim blnSelSpecialGroup
	                                Dim lngSelMinimumRankPosts
	                                Dim blnStartGroup
	                                Dim strLadderGroup
	
	                                'Read in the recordset
	                                intSelGroupID = CInt(rsCommon("Group_ID"))
	                                strSelGroupName = rsCommon("Name")
	                                blnSelSpecialGroup = CBool(rsCommon("Special_rank"))
	                                lngSelMinimumRankPosts = CLng(rsCommon("Minimum_posts"))
	                                blnStartGroup = CBool(rsCommon("Starting_group"))
	                                strLadderGroup = rsCommon("Ladder_Name")
	                                
	                                'Don't allow the member to be made an admin if windows authentication is enabled
	                                If (intSelGroupID > 1 AND blnWindowsAuthentication) OR  (blnWindowsAuthentication = False) Then
	
		                                'Display the selection
		                                Response.Write("<option value=""" & intSelGroupID & """")
	                                
		                                'If this is the group the member is part of then have it slected
		                                If intUsersGroupID = intSelGroupID Then 
		                                	Response.Write(" selected")
		                                'If this is starting group then select it
		                                ElseIf blnStartGroup AND strMode = "new" Then 
		                                	Response.Write(" selected")
		                                End If
		
		                                'Display the end of the select option
		                                If blnSelSpecialGroup Then
		                                        Response.Write(">" & strSelGroupName & "</option>" & vbCrLf)
		                                Else
		                                        Response.Write(">" & strSelGroupName & " - Ladder Group '" & strLadderGroup & "' - Min. Posts " & lngSelMinimumRankPosts & "</option>" & vbCrLf)
		                                End If
		                        End If
	
	                                'Move to the next record
	                                rsCommon.MoveNext
	
	                        Loop
%>
      </select>
    </td>
  </tr>
  <%
		End If
	End If	
%>
  <tr>
    <td width="50%" class="tableRow">Member Title</td>
    <td width="50%" valign="top" class="tableRow"><input name="memTitle" type="text" value="<% = strMemberTitle %>" size="30" maxlength="40" /></td>
  </tr>
  <tr class="tableRow">
    <td>Number of Points</td>
    <td><input type="text" name="points" id="points" size="4" maxlength="7" value="<% = lngMemberPoints %>" /></td>
   </tr>
  <tr>
    <td width="50%" class="tableRow">Number of posts</td>
    <td width="50%" valign="top" class="tableRow"><input name="posts" type="text" value="<% = lngPosts %>" size="4" maxlength="7" /></td>
  </tr>
  <tr class="tableRow">
    <td valign="top">Admin Notes<br /><psan class="smText">These are notes left by Admin and Moderators, about this member, such as any warnings the member has been given (max 250 characters).</span></td>
    <td><textarea name="notes" id="notes" cols="30" rows="4" onKeyDown="characterCounter('notesChars', 'notes');" onKeyUp="characterCounter('notesChars', 'notes');"><% = strAdminNotes %></textarea>
     <br />
     <input size="3" value="0" name="notesChars" id="notesChars" maxlength="3" />
     <input onclick="characterCounter('notesChars', 'notes');" type="button" value="<% = strTxtCharacterCount %>" name="Count" />
   </td>
  </tr>
  <tr>
    <td height="2" colspan="2" align="center" valign="top" class="tableBottomRow">
      <input type="hidden" name="postBack" value="True" />
      <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
      <input type="hidden" name="PF" id="PF" value="<% = lngUserProfileID %>" />
      <input type="submit" name="Submit" id="Submit" value="<% If strMode = "edit" Then Response.Write("Update Member") Else Response.Write("Add New Member") %>" onclick="return CheckForm();" />
      <input type="reset" name="Reset" id="Reset" value="Reset Form" />
    </td>
  </tr>
</table>
</form><%

End If

%>
<br />
<%

'Clean up
Call closeDatabase()

'If the username is already gone diaply an error message pop-up
If blnUsernameOK = False Then
        Response.Write("<script  language=""JavaScript"">")
        Response.Write("alert('Sorry, the Username you requested is already taken.\n\nPlease choose another Username.');")
        Response.Write("</script>")

End If
%>
<!-- #include file="includes/admin_footer_inc.asp" -->
