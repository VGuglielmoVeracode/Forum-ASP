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



Const strTxtTextFormat = "Text Format"
Const strTxtMode = "Mode"
Const strTxtPrompt = "Prompt"
Const strTxtBasic = "Basic"
Const strTxtAddEmailLink = "Add Email Link"
Const strTxtList = "List"
Const strTxtCentre = "Centre"

Const strTxtEnterBoldText = "Enter text you want formatted in Bold"
Const strTxtEnterItalicText = "Enter text you want formatted in Italic"
Const strTxtEnterUnderlineText = "Enter text you want Underlined"
Const strTxtEnterCentredText = "Enter text you want Centred"
Const strTxtEnterHyperlinkText = "Enter the on screen display text for the Hyperlink"
Const strTxtEnterHeperlinkURL = "Enter the URL address to create Hyperlink to"
Const strTxtEnterEmailText = "Enter the on screen display text for the email address"
Const strTxtEnterEmailMailto = "Enter the email address to link to"
Const strTxtEnterImageURL = "Enter the web address of the image"
Const strTxtEnterTypeOfList = "Type of list"
Const strTxtEnterEnter = "Enter"
Const strTxtEnterNumOrBlankList = "for numbered or leave blank for bulleted"
Const strTxtEnterListError = "ERROR! Please enter"
Const strEnterLeaveBlankForEndList = "List item Leave blank to end list"
Const strTxtErrorInsertingObject = "Error inserting object in current location."


Const strTxtFontStyle = "Format"
Const strTxtFontTypes = "Font"
Const strTxtFontSizes ="Size"
Const strTxtEmoticons = "Emoticons"
Const strTxtFontSize = "Font Size"


Const strTxtFontColours ="Font Colours"
Const strTxtBlack = "Black"
Const strTxtWhite = "White"
Const strTxtBlue = "Blue"
Const strTxtRed = "Red"
Const strTxtGreen = "Green"
Const strTxtYellow = "Yellow"
Const strTxtOrange = "Orange"
Const strTxtBrown = "Brown"
Const strTxtMagenta = "Magenta"
Const strTxtCyan = "Cyan"
Const strTxtLimeGreen = "Lime Green"



Const strTxtCut = "Cut"
Const strTxtCopy = "Copy"
Const strTxtPaste = "Paste"
Const strTxtBold = "Bold"
Const strTxtItalic = "Italic"
Const strTxtUnderline = "Underline"
Const strTxtLeftJustify = "Left Justify"
Const strTxtCentrejustify = "Centre Justify"
Const strTxtRightJustify = "Right Justify"
Const strTxtJustify = "Justify"
Const strTxtUnorderedList = "Unordered List"
Const strTxtOutdent = "Outdent"
Const strTxtIndent = "Indent"
Const strTxtAddHyperlink = "Insert Hyperlink"
Const strTxtAddImage = "Insert Image"
Const strTxtJavaScriptEnabled = "JavaScript must be enabled on your web browser in order to use this Rich Text Editor!"
Const strTxtFontColour = "Colour"
Const strTxtstrTxtOrderedList = "Ordered List"
Const strTxtTextColour = "Text Colour"
Const strTxtBackgroundColour = "Background Colour"
Const strTxtUndo = "Undo"
Const strTxtRedo = "Redo"
Const strTxtstrSpellCheck = "Spell Check"
Const strTxtToggleHTMLView = "Toggle HTML View"
Const strTxtAboutRichTextEditor = "About Rich Text Editor"
Const strTxtInsertTable = "Insert Table"
Const strTxtSpecialCharacters = "Special Characters"
Const strTxtPrint = "Print"
Const strTxtImage = "Image"
Const strTxtStrikeThrough = "Strike Through"
Const strTxtSubscript = "Subscript"
Const strTxtSuperscript = "Superscript"
Const strTxtHorizontalRule = "Horizontal Rule"


Const strTxtIeSpellNotDetected = "You need \'ieSpell\' spelling checker installed to use this feature. \nClick OK to go to the \'ieSpell\' download page."
Const strTxtSpellBoundNotDetected = "You need \'SpellBound 0.7.0+\' spelling checker installed to use this feature. \nClick OK to go to the \'SpellBound\' download page."




Const strTxtOK = "OK"
Const strTxtCancel = "Cancel"


Const strTxtImageUpload = "Image Upload"
Const strTxtFileUpload = "File Upload"
Const strTxtUpload = "Upload"
Const strTxtPath = "Path"
Const strTxtFileURL = "File URL"

Const strTxtParentDirectory = "Parent Directory"

Const strTxtImagesMustBeOfTheType = "Images must be of the type"
Const strTxtAndHaveMaximumFileSizeOf = "and have a maximum file size of"
Const strTxtImageOfTheWrongFileType = "The image uploaded is of the wrong file type"
Const strTxtImageFileSizeToLarge = "The image file size is to large at"
Const strTxtMaximumFileSizeMustBe = "Maximum file size must be"
Const strTxtErrorUploadingImage = "Error uploading image!!"
Const strTxtNoImageToUpload = "Please use the \'Browse...\' button to select an image to upload."

Const strTxtFile = "File"
Const strTxtFilesMustBeOfTheType = "Files must be of the type"
Const strTxtFileOfTheWrongFileType = "The file uploaded is of the wrong file type"
Const strTxtFileSizeToLarge = "The file size is to large at"
Const strTxtErrorUploadingFile = "Error uploading file!!"
Const strTxtNoFileToUpload = "Please use the \'Browse...\' button to select a file to upload."


Const strTxtPleaseWaitWhileFileIsUploaded = "Please be patient while the file is being uploaded to the server."
Const strTxtPleaseWaitWhileImageIsUploaded = "Please be patient while the image is being uploaded to the server."


Const strTxtCloseWindow = "Close Window"


Const strTxtPreview = "Preview"
Const strTxtThereIsNothingToPreview = "There is nothing to preview"

Const strResetFormConfirm = "Are you sure you want to reset the form?"
Const strResetWarningFormConfirm = "WARNING: All form data will be lost!!"
Const strResetWarningEditorConfirm = "WARNING: All editor data will be lost!!"


Const strTxtSubmitForm = "Submit Form"
Const strTxtResetForm = "Reset Form"

Const strTxtDisplayMessage = "Display Message"
Const strTxtThereIsNothingToShow = "There is no message to display"


Const strTxtTableProperties = "Table Properties"

Const strTxtImageProperties = "Image Properties"

Const strTxtImageURL = "Image&nbsp;URL"
Const strTxtAlternativeText = "Alt Text"
Const strTxtLayout = "Layout"
Const strTxtAlignment = "Alignment"
Const strTxtBorder = "Border"
Const strTxtSpacing = "Spacing"
Const strTxtHorizontal = "Horizontal"
Const strTxtVertical = "Vertical"

Const strTxtRows = "Rows"
Const strTxtColumns = "Columns"
Const strTxtWidth = "Width"
Const strTxtpixels = "pixels"
Const strTxtCellPad = "Cell Pad"
Const strTxtCellSpace = "Cell Space"

Const strTxtHeight = "Height"


Const strTxtSelectTextToTurnIntoHyperlink = "Please select some text to turn into a hyperlink."

Const strTxtYourBrowserSettingsDoNotPermit = "Your browser settings do not permit the editor to invoke"
Const strTxtPleaseUseKeybordsShortcut = "operations. \nPlease use the keyboard shortcut "
Const strTxtWindowsUsers = "Windows users: "
Const strTxtMacUsers = "Mac users: "


Const strTxtHyperlinkProperties = "Hyperlink Properties"
Const strTxtNoPreviewAvailableForLink = "No Preview Available"
Const strTxtAddress = "Address"
Const strTxtLinkType = "Link Type"
Const strTxtTitle = "Title"
Const strTxtWindow = "Window"
Const strTxtEmail = "Email"
Const strTxtSubject = "Subject"
Const strTxtPleaseWaitWhilePreviewLoaded = "Please be patient while the preview is loaded...."
Const strTxtErrorLoadingPreview = "Error loading preview.\nPlease check the path and name are correct."


Const strTxAttachFileProperties = "Attach File Properties"

Const strTxtNewBlankDoc = "New Blank Document"
Const strTxtOpen = "Open"
Const strTxtSave = "Save"

Const strTxtFileAlreadyExistsRenamedAs = "Either a file of the same name already existed or there was a problem with the file name entered.\nThe file has been saved as"
Const strTxtTheFile = "The file"
Const strTxtHasBeenSaved = "has been saved"


Const strTxtPasteFromWord = "Paste From Word"
Const strTxtPasteFromWordDialog = "This will clean up documents pasted from Word. Please paste inside the following box using the keyboard (Windows users: Ctrl + 'v', MAC users: Apple + 'v') and hit 'OK'."


Const strTxtUploadingFiles = "Uploading Files"
Const strTxtEstimatedTimeLeft = "Estimated time left"
Const strTxtTransferRate = "Transfer rate"
Const strTxtCopied = "copied"
Const strTxtOfFilesUploadedToRemoteServer = "of files uploaded to remote server"

%>