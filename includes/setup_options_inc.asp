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




'******************************************
'***  	  	 Setup options        ****
'******************************************

'Advanced options that are best left as default, except in some circumstances


'Set up application variables prefix
'This can be useful if you are running mutiple installations of Web Wiz Forums on the same site or if you are using free web where you share your application object with others
Const strAppPrefix = "WWF11"


'Set if Application variables are used to improve performance
Const blnUseApplicationVariables = true


'Set Encrypted passwords (ignore unless you don't wish to use Encrypted passwords in your forum)
'This will make your forum vulnerble to hackers if you disable this!!!!!
'This can NOT be changed once your forum is in use!!!
'If you do disable Encrypted Passwords - You will also need to directly edit the database to type in the admin password to the Password field in the " & strDbTable & "Author table at record position 1 also edit both common.asp files to change this variable
Const blnEncryptedPasswords = true 'True = Encrypted Passwords Enabled  -  Flase = Encrypted Passwords Disabled



'Logging
'Web Wiz Forums is able to create log files of actions, activity, and errors. However this WILL effect performance so be careful what you enable logging for.
'Logging requires that the folder storing log files has read, write, and modify permissions for the IUSR account in order for Web Wiz Forums to create and write to log files.
'SECURITY ALERT: MAKE SURE THAT YOU MOVE THE LOG FILE LOCATION TO A FOLDER NOT ACCESSIBLE TO THE PUBLIC.
Const blnLoggingEnabled = False		'Enable logging
Dim strLogFileLocation
strLogFileLocation = Server.MapPath("log_files") 	'Default log file folder, change this to a folder outside your website root if you don't want logs files to be public

Const blnModeratorLogging = True 	'Log the actions of moderators
Const blnErrorLogging = True		'Log error messages
Const blnNewRegistrationLogging = True 	'Log new registrations

Const blnCreatePostLogging = False 	'Log the creating of new topics and posts (Don't enable this on busy forums)
Const blnEditPostLogging = False 	'Log the editing of topics and posts (Don't enable this on busy forums)
Const blnDeletePostLogging = True 	'Log the deletion of topics and posts



'Upload folder path, DO NOT CHANGE as it may break your forums upload tools
Dim strUploadFilePath
strUploadFilePath = "uploads" 'This is the upload folder



'Web Wiz Forums Powered By Links
Const blnWebWizForumsLinks = True

'Toggle between text and image links
Const blnTextLinks = True

%>