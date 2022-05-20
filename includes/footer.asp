<%



'If mobile view and custom Footer enabled
If blnShowMobileHeaderFooter AND blnMobileBrowser Then
	
	'Display custom mobile Footer
	Response.Write(strFooterMobile)

'If custom Footer is enabled	
ElseIf blnShowHeaderFooter Then
	
	'Show custom header
	Response.Write(strFooter)

'Else show standard forum Footer	
Else
%>
</body>
</html><%

End If

%>