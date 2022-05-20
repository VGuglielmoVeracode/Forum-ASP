<link rel="icon" href="favicon.ico" type="image/x-icon" />
<link rel="shortcut icon" href="favicon.ico" type="image/x-icon" />
<meta name="viewport" content="width=device-width">
<meta http-equiv="Content-Script-Type" content="text/javascript" />
<script language="javascript" src="includes/default_javascript_v9.js" type="text/javascript"></script>
<%  

'Anaytics/Stats Tracking code
Response.Write(strStatsTrackingCode)


'If mobile view and custom header enabled
If blnShowMobileHeaderFooter AND blnMobileBrowser Then
	
	'Display custom mobile header
	Response.Write(strHeaderMobile)

'If custom header is enabled	
ElseIf blnShowHeaderFooter Then
	
	'Show custom header
	Response.Write(strHeader)

'Else show standard forum header	
Else
%>
</head>
<body>
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr> 
  <td><% 
   
	'If there is a forum image then dsiplay it
	If NOT strTitleImage = "" Then Response.Write("<a href=""" & strWebsiteURL & """><img src=""" & strTitleImage & """ border=""0"" alt=""" & strWebsiteName & " " & strTxtHomepage & """ title=""" & strWebsiteName & " " & strTxtHomepage & """ /></a>")

%></td>
 </tr>
</table><%

End If

%>