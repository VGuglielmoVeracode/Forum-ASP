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




'Function to send an e-mail
Function SendMail(ByVal strEmailBodyMessage, ByVal strRecipientName, ByVal strRecipientEmailAddress, ByVal strFromEmailName, ByVal strFromEmailAddress, ByVal strSubject, strMailComponent, blnHTML)

	'Dimension variables
	Dim objCDOSYSMail		'Holds the CDOSYS mail object
	Dim objCDOMail			'Holds the CDONTS mail object
	Dim objJMail			'Holds the Jmail object
	Dim objAspEmail			'Holds the Persits AspEmail email object
	Dim objAspMail			'Holds the Server Objects AspMail email object
	Dim strEmailBodyAppendMessage	'Holds the appended email message
	
	
	'If we are in demo mode we don't want to send emails so exit function
	If blnDemoMode Then 
		SendMail = False
		Exit Function
	End If
	
	
	'Set error trapping
	On Error Resume Next
	
	
	
	'Remove unwanted cahracters that may course the email component to throw an exception
	'or be used by a spammer to send out BCC spam emails using a malformed form entry
	
	strSubject = Trim(Mid(Replace(strSubject, vbCrLf, ""), 1, 100))
	
	strRecipientName = Trim(Mid(strRecipientName, 1, 35))
	strFromEmailName = Trim(Mid(strFromEmailName, 1, 35))
	strRecipientEmailAddress = Trim(Mid(strRecipientEmailAddress, 1, 50))
	strFromEmailAddress = Trim(Mid(strFromEmailAddress, 1, 50))
	
	strRecipientName = Replace(strRecipientName, vbCrLf, "", 1, -1, 1)
	strFromEmailName = Replace(strFromEmailName, vbCrLf, "", 1, -1, 1)
	strRecipientEmailAddress = Replace(strRecipientEmailAddress, vbCrLf, "", 1, -1, 1)
	strFromEmailAddress = Replace(strFromEmailAddress, vbCrLf, "", 1, -1, 1)
	
	strRecipientName = Replace(strRecipientName, ",", "", 1, -1, 1)
	strFromEmailName = Replace(strFromEmailName, ",", "", 1, -1, 1)
	strRecipientEmailAddress = Replace(strRecipientEmailAddress, ",", "", 1, -1, 1)
	strFromEmailAddress = Replace(strFromEmailAddress, ",", "", 1, -1, 1)
	
	strRecipientName = Replace(strRecipientName, ";", "", 1, -1, 1)
	strFromEmailName = Replace(strFromEmailName, ";", "", 1, -1, 1)
	strRecipientEmailAddress = Replace(strRecipientEmailAddress, ";", "", 1, -1, 1)
	strFromEmailAddress = Replace(strFromEmailAddress, ";", "", 1, -1, 1)
	
	strRecipientName = Replace(strRecipientName, ":", "", 1, -1, 1)
	strFromEmailName = Replace(strFromEmailName, ":", "", 1, -1, 1)
	strRecipientEmailAddress = Replace(strRecipientEmailAddress, ":", "", 1, -1, 1)
	strFromEmailAddress = Replace(strFromEmailAddress, ":", "", 1, -1, 1)
	
	strRecipientName = Replace(strRecipientName, "<", "", 1, -1, 1)
	strFromEmailName = Replace(strFromEmailName, "<", "", 1, -1, 1)
	strRecipientEmailAddress = Replace(strRecipientEmailAddress, "<", "", 1, -1, 1)
	strFromEmailAddress = Replace(strFromEmailAddress, "<", "", 1, -1, 1)
	
	strRecipientName = Replace(strRecipientName, ">", "", 1, -1, 1)
	strFromEmailName = Replace(strFromEmailName, ">", "", 1, -1, 1)
	strRecipientEmailAddress = Replace(strRecipientEmailAddress, ">", "", 1, -1, 1)
	strFromEmailAddress = Replace(strFromEmailAddress, ">", "", 1, -1, 1)
	
	


	'Check the email body doesn't already have Web Wiz Forums
	If blnLCode Then
		
		
		'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******

		'If HTML format then make an HTML link
		If blnHTML = True Then
			strEmailBodyAppendMessage = "<br /><br /><br /><hr />Software provided by <a href=""http://www.webwizforums.com"">Web Wiz Forums&reg;</a> version " & strVersion & " - http://www.webwizforums.com"
		'Else do a text link
		Else
			strEmailBodyAppendMessage = VbCrLf & VbCrLf & "---------------------------------------------------------------------------------------"  & _
			vbCrLf & "Software provided by Web Wiz Forums(TM) version " & strVersion& " - http://www.webwizforums.com"
		End If
		
		'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
	End If
	
	
	'Check to see if an error has occurred
	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then Call errorMsg("An error has occurred while sending an email.", "send_mail_header", "functions_send_mail.asp")



	'******************************************
	'***	        Mail components        ****
	'******************************************

	'Select which email component to use
	Select Case strMailComponent



		'******************************************
		'***	  MS CDOSYS mail component     ****
		'******************************************

		'CDOSYS mail component
		Case "CDOSYS", "CDOSYSp", "CDOSYSSSL"

			'Dimension variables
			Dim objCDOSYSCon
			Dim intSendUsing
			
			'Port or pick up directory (1=pick up directory(localhost) 2=port(network))
			If strMailComponent = "CDOSYSp" Then
				intSendUsing = 1
			Else
				intSendUsing = 2
			End If

			'Create the e-mail server object
			Set objCDOSYSMail = Server.CreateObject("CDO.Message")
		    Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")
		    	
		    'Check to see if an error has occurred
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then Call errorMsg("An error has occurred while sending an email.<br />Please check that the CDOSYS email component is installed on the server.", "create_CDOSYS_object", "functions_send_mail.asp")
		

		    	'Set and update fields properties
		    	With objCDOSYSCon
		    		
		    		'Use SMTP Server authentication if required
		    		If strMailServerUser <> "" AND strMailServerPass <> "" Then 
			    		' Specify the authentication mechanism to basic (clear-text) authentication cdoBasic = 1
			        	.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
			        	
			        	'SMTP Server username
			        	.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = strMailServerUser
			        	
			        	'SMTP Server password
			        	.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strMailServerPass
			        End If
		        	
		        	
		        	'Out going SMTP server
		        	.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strMailServer
		        	
		        	'SMTP port
		        	.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport")  = lngMailServerPort
		        	
		        	'SSL
		        	If strMailComponent = "CDOSYSSSL" Then .Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = 1
		        	
		        	'CDO Port (1=localhost 2=network)
		        	.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = intSendUsing
		        	
		        	'Set CDO pickup directory if using localhost (CDO Port 1)
		        	If intSendUsing = 1 Then
		        		'CDO pickup directory (used for localhost service)
		        		.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = "c:\inetpub\mailroot\pickup" 
		        	End If
		        	
		        	'Timeout
		        	.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
	        		.Fields.Update
	        	End With

			'Update the CDOSYS Configuration
			Set objCDOSYSMail.Configuration = objCDOSYSCon

			With objCDOSYSMail
				'Who the e-mail is from
				.From = strFromEmailName & " <" & strFromEmailAddress & ">"

				'Who the e-mail is sent to
				.To = strRecipientName & " <" & strRecipientEmailAddress & ">"
				
				'Set the charcater encoding for the email
				.BodyPart.Charset = strPageEncoding 

				'The subject of the e-mail
				.Subject = strSubject

				'Set the e-mail body format (HTMLBody=HTML TextBody=Plain)
				If blnHTML = True Then
				 	.HTMLBody = strEmailBodyMessage & strEmailBodyAppendMessage
				Else
					.TextBody = strEmailBodyMessage & strEmailBodyAppendMessage
				End If

				'Send the e-mail
				If NOT strMailServer = "" Then .Send
			End with

			'Close the server mail object
			Set objCDOSYSMail = Nothing
			Set objCDOSYSCon = Nothing




		'******************************************
		'***  	  MS CDONTS mail component     ****
		'******************************************

		'CDONTS mail component
		Case "CDONTS"

			'Create the e-mail server object
			Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
			
			'Check to see if an error has occurred
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then Call errorMsg("An error has occurred while sending an email.<br />Please check that the CDONTS email component is installed on the server.", "create_CDONTS_object", "functions_send_mail.asp")
		

			With objCDOMail
				'Who the e-mail is from
				.From = strFromEmailName & " <" & strFromEmailAddress & ">"

				'Who the e-mail is sent to
				.To = strRecipientName & " <" & strRecipientEmailAddress & ">"

				'The subject of the e-mail
				.Subject = strSubject

				'The main body of the e-amil
				.Body = strEmailBodyMessage & strEmailBodyAppendMessage

				'Set the e-mail body format (0=HTML 1=Text)
				If blnHTML = True Then
					.BodyFormat = 0
				Else
					.BodyFormat = 1
				End If

				'Set the mail format (0=MIME 1=Text)
				.MailFormat = 0

				'Importance of the e-mail (0=Low, 1=Normal, 2=High)
				.Importance = 1

				'Send the e-mail
				.Send
			End With

			'Close the server mail object
			Set objCDOMail = Nothing




		'******************************************
		'***  	  w3 JMail mail component      ****
		'******************************************

		'JMail component
		Case "Jmail"

			'Create the e-mail server object
			Set objJMail = Server.CreateObject("JMail.SMTPMail")
			
			'Check to see if an error has occurred
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then Call errorMsg("An error has occurred while sending an email.<br />Please check that the JMail email component is installed on the server.", "create_JMail_3_object", "functions_send_mail.asp")
		

			With objJMail
				'Out going SMTP mail server address
				.ServerAddress = strMailServer & ":" & lngMailServerPort

				'Who the e-mail is from
				.Sender = strFromEmailAddress
				.SenderName = strFromEmailName

				'Who the e-mail is sent to
				.AddRecipient strRecipientEmailAddress
				
				'Set the charcater set for the email
				.Charset = strPageEncoding

				'The subject of the e-mail
				.Subject = strSubject

				'Set the e-mail body format (BodyHTML=HTML Body=Text)
				If blnHTML = True Then
					.HTMLBody = strEmailBodyMessage & strEmailBodyAppendMessage
				Else
					.Body = strEmailBodyMessage & strEmailBodyAppendMessage
				End If

				'Importance of the e-mail
				.Priority = 3

				'Send the e-mail
				If NOT strMailServer = "" Then .Execute
			End With

			'Close the server mail object
			Set objJMail = Nothing
		
		
		
		
		'******************************************
		'***   w3 JMail ver.4+ mail component  ****
		'******************************************

		'JMail ver.4+ component (this version allows authentication)
		Case "Jmail4"

			'Create the e-mail server object
			Set objJMail = Server.CreateObject("Jmail.Message")
			
			'Check to see if an error has occurred
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then Call errorMsg("An error has occurred while sending an email.<br />Please check that JMail 4 or above email component is installed on the server.", "create_JMail_4_object", "functions_send_mail.asp")
		

			With objJMail
			
				'Send SMTP Server authentication data
    				If NOT strMailServerUser = "" Then .MailServerUserName = strMailServerUser
    				If NOT strMailServerPass = "" Then .MailServerPassword = strMailServerPass

				'Who the e-mail is from
				.From = strFromEmailAddress
				.FromName = strFromEmailName

				'Who the e-mail is sent to
				.AddRecipient strRecipientEmailAddress, strRecipientName
				
				'Do not encode as text is already encoded
				.EnableCharsetTranslation = False
				
				'Set the charcater set for the email
				.Charset = strPageEncoding

				'The subject of the e-mail
				.Subject = .EncodeHeader(strSubject, strPageEncoding)

				'Set the e-mail body format (BodyHTML=HTML Body=Text)
				If blnHTML = True Then
					.ContentType = "text/html"
					.HTMLBody = strEmailBodyMessage & strEmailBodyAppendMessage
				Else
					.Body = strEmailBodyMessage & strEmailBodyAppendMessage
				End If

				'Importance of the e-mail
				.Priority = 3

				'Send the e-mail
				If NOT strMailServer = "" Then .Send(strMailServer & ":" & lngMailServerPort)
			End With

			'Close the server mail object
			Set objJMail = Nothing




		'******************************************
		'*** Persits AspEmail mail component   ****
		'******************************************

		'AspEmail component
		Case "AspEmail", "AspEmailTLS"

			'Create the e-mail server object
			Set objAspEmail = Server.CreateObject("Persits.MailSender")
			
			'Check to see if an error has occurred
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then Call errorMsg("An error has occurred while sending an email.<br />Please check that the AspEmail email component is installed on the server.", "create_AspEmail_object", "functions_send_mail.asp")
		

			With objAspEmail
				'Out going SMTP mail server address
				.Host = strMailServer
				.Port = lngMailServerPort
				
				'TLS SSL
				If strMailComponent = "AspEmailTLS" Then .TLS = 1
				
				'Use SMTP Server authentication if required
		    		If strMailServerUser <> "" AND strMailServerPass <> "" Then 
					'SMTP server username and password
					.Username = strMailServerUser
					.Password = strMailServerPass
				End If

				'Who the e-mail is from
				.From = strFromEmailAddress
				.FromName = strFromEmailName

				'Who the e-mail is sent to
				.AddAddress strRecipientEmailAddress
				
				'Set the charcater set for the email
				.Charset = strPageEncoding
				.ContentTransferEncoding = "Quoted-Printable"

				'The subject of the e-mail
				.Subject = strSubject

				'Set the e-mail body format (BodyHTML=HTML Body=Text)
				If blnHTML = True Then
					.IsHTML = True
				End If

				'The main body of the e-mail
				.Body = strEmailBodyMessage & strEmailBodyAppendMessage

				'Send the e-mail
				If NOT strMailServer = "" Then .Send
			End With

			'Close the server mail object
			Set objAspEmail = Nothing




		'********************************************
		'*** ServerObjects AspMail mail component ***
		'********************************************

		'AspMail component
		Case "AspMail"

		   	'Create the e-mail server object
		   	Set objAspMail = Server.CreateObject("SMTPsvg.Mailer")
		   	
		   	'Check to see if an error has occurred
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then Call errorMsg("An error has occurred while sending an email.<br />Please check that the AspMail email component is installed on the server.", "create_AspMail_object", "functions_send_mail.asp")
		

		   	With objAspMail
			   	'Out going SMTP mail server address
			   	.RemoteHost = strMailServer & ":" & lngMailServerPort

			   	'Who the e-mail is from
			   	.FromAddress = strFromEmailAddress
			   	.FromName = strFromEmailName

			   	'Who the e-mail is sent to
			   	.AddRecipient " ", strRecipientEmailAddress

			   	'The subject of the e-mail
			   	.Subject = strSubject

			   	'Set the e-mail body format (BodyHTML=HTML Body=Text)
			   	If blnHTML = True Then
			    		.ContentType = "text/HTML"
			   	End If

			   	'The main body of the e-mail
			   	.BodyText = strEmailBodyMessage & strEmailBodyAppendMessage

			   	'Send the e-mail
			   	If NOT strMailServer = "" Then .SendMail
			   End With

		   	'Close the server mail object
		   	Set objAspMail = Nothing
	End Select
	
	
	'Check to see if an error has occurred
	If Err.Number <> 0 Then 
		
		'If logging is enabled write the error message to the log file
		If blnLoggingEnabled Then Call logAction(strLoggedInUsername, "ERROR - File: " & strFileName & " - Error Details: err_" & strDatabaseType & "_" & strErrCode & " - " & Err.Source & " - " & Err.Description)
	
		'Place error message into email error variable
		strEmailErrorMessage = "<br />" & Err.Source & "<br />" & Err.Description & "<br /><br />"
		
		'Set the returned value of the function to true
		SendMail = False
		
		'Display an error message to screen (disabled to prevent users seeing error messages)
		'Call errorMsg("An error has occurred while sending an email.", "send_mail_footer", "functions_send_mail.asp")
		
	'Else the email has been sucessfully sent	
	Else
		'Set the returned value of the function to true
		SendMail = True
	End If

	'Disable error trapping
	On Error goto 0

End Function
%>