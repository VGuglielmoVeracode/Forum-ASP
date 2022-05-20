<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz Rich Text Editor(TM)
'**  http://www.richtexteditor.org
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

Dim blnAbout

'Initiliase variables
Const strRTEversion = "4.18wwf"
blnAbout = false



'The following enables and disables functions of the Rich Text Editor

'Enable and dsiable basic functions of the editor change the following to true of false
'***************************************************************************
Const blnNew = false
Const blnBold = true
Const blnUnderline = true
Const blnItalic = true
Const blnFontStyle = false
Const blnFontType = true
Const blnFontSize = true
Const blnTextColour = true
Const blnTextBackgroundColour = false
Const blnCut = true
Const blnCopy = true
Const blnPaste = true
Const blnWordPaste = true
Const blnUndo = false
Const blnRedo = false
Const blnLeftJustify = true
Const blnCentre = true
Const blnRightJustify = true
Const blnFullJustify = false
Const blnOrderList = true
Const blnUnOrderList = true
Const blnOutdent = true
Const blnIndent = true
Const blnAddHyperlink = true
Const blnAddImage = true
Const blnInsertTable = false
Const blnSpecialCharacters = true
Const blnPrint = false
Const blnStrikeThrough = true
Const blnSubscript = false
Const blnSuperscript = false
Const blnHorizontalRule = false
Const blnPreview = false
'***************************************************************************

'BB Code extras
'***************************************************************************
Const blnQuoteBlock = true
Const blnCodeBlock = true
'***************************************************************************


'Advanced controls
'***************************************************************************
Const blnAdvAdddHyperlink = true 'Advanced hyperlink control
Const blnAdvAddImage = true 	'Advanced image control requires File System Object (FSO)
Const blnHTMLView = false	'Allows the user to view the HTML code, you may need to dsiable this for extra security
Const blnSpellCheck = true	'Requires IEspell for Ineternet Explorer or SpellBound for Mozilla
Const blnUseCSS = false		'Enable CSS (Cascading Style Sheets) in Mozilla
Const blnNoIEdblLine = true	'Prevent IE's standard double line spacing when the 'ENTER' key is pressed
'***************************************************************************



'Using full URL path for images and links
'***************************************************************************
'If you are submitting the RTE content to a file outside of the RTE folder you may find that some of the relative
'paths for things like images stored on the server are incorrect (ie. href="my_documents/myPicture.jpg")
'The following can be used to change those relative server paths to full URL's so that if the submitted content is 
'displayed on a page out side of the RTE files the paths to images etc. still work

Const blnUseFullURLpath = false
Const strFullURLpathToRTEfiles = "" 'Type in the full URL to the RTE folder eg. "http://www.myweb.com/RTE/"

'***************************************************************************
%>