<% @ Language=VBScript %>
<%
'***** WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******


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






'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******



'Function to send convert hex to binary output for browser
Private Function displayImage(strHexImage)

	'Dimension variables
	Dim sarryHex		'Holds each hex value in the image
	Dim lngHexArrayPosition	'Holds the array position

	'Place the Hex image into an array
	sarryHex = split(strHexImage, " ")

	'Loop through the array to convert to binary
	For lngHexArrayPosition = 0 to CLng(UBound(sarryHex))

		'Write the binary output
 		Response.BinaryWrite ChrB(CLng("&H" & sarryHex(lngHexArrayPosition)))
	Next

End Function





Dim strHexImage

'Set so the CAPTCHA image is not cached
Response.CacheControl = "Store"
Response.AddHeader "pragma","cache"
Response.AddHeader "cache-control","public"


'Set the content type for the CAPTCHA image
Response.ContentType = "image/gif"


'***** WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******


strHexImage = "" & _
"47 49 46 38 39 61 ad 00 26 00 f7 00 00 33 33 33 00 00 f6 0d 0e da cc 00 00 00 99 00 99 99 99 ff eb e6 e6 ca 22 7d 7d 7d e6 e6 e6 73 73 73 4c 4c 4c e8 8c 00 99 99 ff bd bd bd a5 a5 a5 e2 51 51 " & _
"a5 d0 8f 6d 6d e9 f2 bd b7 00 e6 00 ff 99 99 ff 79 79 b5 b5 b5 87 c4 75 c8 2c 39 5a ff 59 cc cc cc 33 99 33 51 51 ed 96 dc 77 e4 e0 29 66 66 66 ff f6 9f 18 ff 26 bc bc ff 9a 97 df 41 41 41 ff " & _
"ec 52 ec e6 c8 00 cc 00 41 44 f6 ac ac ac ff ff 00 bc 26 27 8c 8c 8c d8 d8 d8 f2 a8 2e 59 59 59 75 74 db e0 fe ee a8 ff a7 cc 66 66 ff cc 00 ed 2c 37 34 35 f7 68 69 dd d0 df b4 ff f6 8f 85 85 " & _
"e9 9f ff ac a6 de 92 ff f7 00 00 b5 00 ff ff b6 0c e5 15 ff 99 00 ff 00 00 f2 d4 d0 f9 dd 65 5d 5f dc af af e3 ff ff 33 b6 b3 c7 ff de 00 d8 db ed ff 66 66 ff 19 19 c6 f8 c7 c5 c5 c5 21 21 e5 " & _
"3c ff 3b cc cc ff ff ff 66 ff ff ff 0d cd 0d ff 4b 4b ea fa f6 f1 db d6 00 f7 00 2f b9 2f 64 64 ea ff a4 a4 0d a1 0d e7 c3 24 0d b7 0d b9 e9 b1 84 84 84 ff ff df ed 00 00 e8 af 06 e8 8b 8b 66 " & _
"66 ff ff ff 1a ff 42 42 de de de 17 17 e5 8c 8c ff e1 31 31 00 d6 00 80 ff 7f ff c0 00 f2 c4 be 33 ff 33 ff ff 47 0a 8a 0a ff 66 66 b1 d9 aa be bb cc 66 ff 66 ff e5 00 ff 0f 0f ff 21 21 ff f6 " & _
"71 ff ff cc ad ad ff e5 e5 f9 ff e2 4e ff f6 bc 43 d0 49 d6 f6 d6 ff cc cc f1 ae a7 4c 9f 4c 0f ff 1a ff 58 58 3f 3f ff a7 a7 ef 24 24 ff ff ff 54 d6 d6 f9 f7 f7 f7 53 53 ff fe f6 f4 08 ff 08 " & _
"0f 0f e5 7a 7a ff 66 66 ff ff b4 00 a9 ac e1 99 99 99 ff 33 33 ff ff 0f ef ef ef 83 83 ff ff 8d 8d b5 b1 ce ff f6 ae be df a1 e6 e4 c8 00 bd 00 e7 9a 43 ef ae 2b 33 33 ff 80 7e d9 8d cf 7b 00 " & _
"de 00 cc 33 33 ff 84 84 e8 a3 06 e7 2d 3c c2 c2 fc fe 08 08 ec bf 10 70 fc 70 e1 70 70 fc dd dd ff ee 00 df df fe e6 00 00 ff d6 00 7d ff 8b f5 aa a1 1a 1a fe ff 99 99 1c e1 25 b6 b6 ff c5 c3 " & _
"d3 ff c4 c4 67 69 f4 25 28 f6 e2 5e 5e 0d c7 0d ab dd 8f ac ac f6 ff b5 b5 ff d6 d7 ff 52 51 99 99 ff 00 ef 00 00 ac 00 36 a8 36 5b 5b ff a4 a4 ff 5a ff 68 0a 0d f8 db 00 00 49 4c f6 ff 29 29 " & _
"66 66 66 00 ff 00 ff 3b 3b 4a 4a ff d8 2f 3c e6 d8 22 c4 ff d0 ff fe 86 0e f5 18 ff 72 72 a8 a7 e1 da e0 be ee ee fe c8 df ab ff ff e8 22 ff 22 3c 3c ff de de f5 f7 00 00 90 d0 7e dd 2f 3c e6 " & _
"e0 36 ff f5 55 b5 df 98 ff ab 00 ff ff 3a 0d ac 0d 17 1b f7 35 ff 41 ff f6 cb 17 ff 17 b2 b2 e2 8d fc 8d 42 ff 42 ff ff 73 73 73 ff 07 d2 07 00 c4 00 ff ff 99 f5 e5 e1 d5 d5 ff ff ff a7 ff b1 " & _
"ab cc ff cc f5 a1 30 36 b5 36 ff ff ff 21 f9 04 01 07 00 ff 00 2c 00 00 00 00 ad 00 26 00 00 08 ff 00 ff 09 1c 48 b0 a0 c1 83 08 13 2a 5c c8 b0 a1 c3 87 10 23 4a 9c 48 b1 a2 c5 8b 18 33 6a fc " & _
"97 e9 89 83 0b 20 43 8a 1c 49 b2 a4 c9 93 28 53 aa 5c c9 b2 a5 cb 97 20 1d 3c 59 e8 a2 85 82 07 2a 60 ea dc c9 b3 a7 cf 9f 26 55 14 b8 99 e0 a0 83 02 1b a8 28 5d ca b4 a9 d3 a7 50 a3 4a 9d 4a " & _
"b5 aa d5 ab 58 b3 3e 6d 74 74 26 c1 0b 2d 32 69 1d ab 75 83 d9 0d 2e c8 42 75 b1 41 2c d4 0d 09 d4 36 cd 94 56 ae d4 04 61 bc fe 4b d0 a2 91 d2 46 dd fc 08 f6 23 8e 91 61 46 52 12 af 92 f5 c8 " & _
"00 ac c7 c4 22 0f 22 66 b7 29 00 10 98 5b 54 5e 0a 62 01 88 12 17 9e 22 50 70 01 c1 e6 0d 20 36 47 6d 14 a6 e8 bf 07 4f 96 fa e9 d6 4d 16 e2 55 23 fa 2c 6b d0 60 8d a4 5e 24 f0 39 1e c4 6b 18 " & _
"97 0a 9b ca a8 56 0a c0 69 82 a4 54 12 3c 4f 50 37 a9 74 17 69 b1 33 dd 00 43 a9 8b 12 4a 9f 87 ff 87 f1 e0 01 0c b4 de 95 26 cd 84 9e 0a dc b4 ec dd aa 07 41 9d 0a dd f0 71 37 3b 50 b1 b7 80 " & _
"5f a5 e2 04 96 8f 14 ab e8 a2 5b 03 9a 68 42 8f 19 cf 90 82 85 01 c4 14 77 dc 26 da bc b2 1c 15 00 9c 25 d6 05 30 b4 00 43 26 05 2c 50 80 03 a9 61 48 45 01 20 b4 b0 80 02 2d 94 00 9d 7d 0b 3c " & _
"50 17 15 1c 7a c8 de 02 08 20 20 e2 02 54 7c 17 dd 02 2e 88 98 22 86 08 38 e0 02 0c 05 c0 30 e3 06 25 28 a0 80 69 25 c4 a5 80 03 aa 35 12 d6 06 2a 30 c5 88 2c 03 8e a0 cb 32 cb ac e1 db 82 ca " & _
"a4 40 0a 11 44 48 58 01 2a 16 30 61 e1 72 97 61 96 56 94 54 20 f0 40 01 9a a1 c6 dc 89 9a e1 79 62 01 4c 65 f2 80 02 22 52 41 a7 9d 54 80 60 56 6a 4a da e4 c0 03 2d 48 b7 41 0b a9 01 20 16 02 " & _
"48 0d b5 d4 06 3c 52 01 24 a6 99 80 b7 9c 0a 58 66 b9 94 61 52 e4 b6 db 1a 9a 48 42 26 35 37 c4 ff 90 a6 71 9b b0 c9 44 31 bf 5c d8 1c 53 bb 16 e0 2b a0 7a 9a e8 eb 9f c4 aa 07 5d 26 40 f6 0a " & _
"a8 a2 7a de a9 24 02 4a 06 d9 02 a5 26 26 8a 69 01 a1 cd a7 94 a2 41 96 76 a1 4c e5 31 b5 0a 6e 07 b2 2a 89 19 66 38 42 0d 22 be c4 00 c7 3e c8 a1 a2 0d 13 87 a0 01 81 ae 4d 2d 90 14 08 17 0c " & _
"cb 5d 74 cd 0d 2b 30 a0 ea e9 1b 5d 94 06 f3 9b e8 a2 d1 c1 60 5a 67 c4 3e 50 a9 52 98 f6 48 65 c1 99 84 2a 16 0c 0b 5c ac da 05 1e 79 4c 05 b9 ab 2a 88 ae ba 88 8c c2 0e 0e 13 20 67 c1 bc c5 " & _
"a0 01 cd bd 70 36 75 64 87 c5 82 d0 61 c0 80 0e cc d4 03 0b 00 b0 40 68 37 6b b6 70 02 25 68 06 43 68 08 68 76 e4 93 dd ed 9a 09 66 4a 6e ea d9 d2 4a 3d 20 ea 72 20 3b 20 b2 aa bd b9 6a 86 32 " & _
"ea 86 33 8a 22 cd e0 40 c8 26 2f 1f 12 f3 25 4d 94 73 e1 dc 74 ab e5 ed 85 5d 8b 2c 4c 6f 26 93 ff 4d 8d d9 8a dc d2 4c 00 46 10 d2 e6 21 56 a0 71 89 1e ac b0 51 f7 e3 90 4f 75 c1 02 f9 71 1d " & _
"32 53 89 ec 20 81 04 5b 6c d1 41 07 29 dc 00 05 14 6a 44 22 40 0c b6 d0 6b c5 34 d0 e8 91 c7 10 8e 47 2e fb ec 90 e7 cd 94 38 bb d8 a1 bb 1d 49 f4 ce 09 3c f0 54 22 fc 12 5c 14 b3 3a 34 4d b0 " & _
"32 c4 38 b1 2b 45 a9 69 4a 4d 2d e7 52 cf 47 75 01 66 52 5d 0f c2 8c 94 82 f0 c0 a6 98 a1 58 22 55 98 95 6f 3e ed 58 d9 be 94 18 df 88 c1 07 10 fa dc 83 cd 14 7b 88 d1 54 05 d3 5c 82 bc f2 e3 " & _
"8c 71 0a 53 08 00 c0 ae dc 23 40 00 7c 4f 29 25 00 80 02 a2 52 00 01 4a 65 03 02 cc 96 a1 04 d8 1d a5 34 10 00 17 b4 4a 01 37 e8 40 f4 59 45 7d 4a 71 87 fb 80 20 bf 29 2c 02 09 e6 38 41 53 36 " & _
"d1 ba e4 2d 6f 0c fe fb 99 00 a1 73 41 05 46 4f 80 04 7b 4a 06 57 23 40 a3 25 60 83 9c 01 c0 d2 ff b0 57 15 01 8a e8 57 c3 f2 60 55 40 48 05 40 00 22 7e f3 30 a1 3a ce f0 01 15 32 05 15 4d c8 " & _
"03 ff 60 e8 8c ff 6d 4a 80 07 04 41 01 3b 05 c1 0c 31 b0 83 51 09 5a 89 2e b0 41 e8 24 10 7a 57 11 e0 f8 94 98 be cb 2d a5 13 fa c0 46 14 e5 80 84 33 ac c0 1a 56 5c 8a 36 86 40 c8 fe 8d c1 19 " & _
"03 60 81 65 00 60 34 01 2a 40 80 71 b9 a0 5f 1c 10 34 05 e6 e7 82 47 02 40 09 72 b8 94 47 ee 2a 80 9f c1 20 c0 0c a8 bd 44 99 2f 33 8b 9c 63 53 80 e6 c8 ba 5c af 05 6c 4c da 90 94 52 4a 17 60 " & _
"26 13 01 04 00 02 6e 24 b4 6c 51 52 80 25 d8 e5 52 98 18 82 f9 2d 82 8f 98 c0 84 0f 00 79 3f 1b d8 40 15 aa 28 87 34 33 40 83 a6 c0 e0 32 04 04 80 0b 04 48 a5 47 76 c7 01 1c 2c 81 58 2e 98 c0 " & _
"02 1a 6d 29 0f 10 60 5a 82 36 14 1b 82 53 9b 19 e4 a0 1c 53 09 95 5c b6 f1 44 42 2b 60 18 b0 69 ff 41 07 96 f1 9a 05 14 e3 06 13 f0 ce 0d 2e 90 96 76 54 8a 0e a6 20 87 29 62 62 05 3e 88 c5 01 " & _
"3e 71 05 41 58 d4 09 60 88 00 39 c8 51 8a 8e 96 02 03 75 68 4a 2e a9 90 4e f0 5c 13 50 41 83 92 10 5d e0 80 37 e2 73 a5 99 04 40 e5 b2 e9 80 4c cc b0 81 e0 c9 a0 4e 7f 55 ce 03 2e 45 9e 1d 7c " & _
"27 0c 3e 92 c0 9c 02 b3 00 98 e2 e7 4b b3 59 02 07 24 a0 92 30 48 40 3a 31 28 c6 83 e2 d4 2d 4c c4 06 1f cf a0 cc 88 e2 e1 00 dc 90 01 3f 66 10 0f 57 68 a0 1d 22 30 84 5a b3 91 8d 20 04 a2 29 " & _
"6c d4 e6 23 17 08 4a 9b 1a b0 8c d0 29 e9 4b ab 03 c6 45 8e 48 80 8d 28 63 02 1e 99 9a 1d 66 cd 91 4e 01 ea ae 1e 29 4e a5 bc b3 5f dc ec e7 f8 32 58 46 a3 b5 00 92 cc a1 aa 10 a9 d4 88 99 32 " & _
"71 0f 48 48 a6 0f 7c 80 07 3c d0 c2 0b 39 b8 c6 0c dc 70 07 79 80 03 12 d2 c8 c2 31 8e 61 0a 53 ff a0 40 0b 36 8b 60 02 bf 17 d7 32 ba e0 82 bf ca 25 88 d0 88 21 51 32 e5 9a 29 12 62 66 27 c7 " & _
"c8 a5 aa c7 88 ff e1 95 d0 90 48 30 97 fe 14 83 86 6d e0 64 fd 19 59 e7 16 d7 57 06 95 20 8d 12 4a 85 29 f8 31 a2 b1 50 82 12 6a e0 05 6f f0 80 b5 ae 85 44 16 64 4b 01 53 b4 a1 0d f6 f0 47 62 " & _
"15 a8 ce 51 8e b4 86 1c dc 80 61 bf 2b 52 4d 5e d3 68 62 cc 65 68 76 18 2a 4d ce 48 ba aa cc 2c 27 05 98 d4 01 66 97 bb 66 f4 2e 0e 1b 21 d0 80 62 95 bc e7 50 66 2c 4a 4b 0b 5a d4 40 14 9e a8 " & _
"85 06 a2 f0 da f9 52 a0 be b6 45 01 28 f4 db 94 0e 0f b0 a8 d8 bc e0 59 ce d2 88 01 e3 70 95 1b bc d8 65 0b 18 c9 0e 0a 54 bc 10 7e 8a 75 25 7c e1 01 0e 39 9b 34 44 e3 8f 7b 94 a4 02 1e 90 89 " & _
"26 88 c5 88 4d 5b 83 1a bc 41 14 e8 60 46 1c de 21 df 2c bc d8 be f8 b5 c7 0f 92 e1 94 21 2b d5 ff 93 cd 7d 67 75 1e 90 94 0b 5e ec b1 b9 25 b2 63 0b 28 aa 0c 0e 19 8e fb 8d 30 15 18 8b 55 6e " & _
"36 79 29 00 85 72 3f 07 88 c3 a3 40 e7 87 c6 bd 80 03 36 e0 53 2a 98 80 c4 26 f6 32 25 5e 10 8c 76 c0 56 b6 b4 b5 2f 0a 64 0c 0a 64 b0 19 ae 05 cc e1 54 01 90 ad 04 46 2b 68 dd 21 a7 d7 8a 1a " & _
"dd eb 0a 6d 29 8d 08 e8 a2 b7 29 c7 f3 25 d9 29 71 1d aa d6 6e 6d d8 a9 66 aa 80 8a d6 30 06 5d 1d 97 32 66 8b 54 20 8a ae 09 94 50 62 2f bf 81 12 e9 e8 47 30 44 10 db 63 bc f8 be a3 06 45 a9 " & _
"4d ed 9c 02 be a8 8c 32 a5 a5 3c eb 2c c4 72 4e b9 c6 14 66 4a 25 09 a6 63 c5 d2 53 34 e1 4c 8b 61 21 5d c0 44 e3 75 d1 d7 c5 96 3c 2b d8 e3 4c bc 26 36 4a f9 43 b5 df 70 ed 74 64 db 03 86 00 " & _
"75 6d db 80 02 7b a8 f9 07 c8 20 00 07 9e 52 be 1a 13 d1 b1 00 55 14 2d 53 94 00 31 c6 08 2a 12 9e 03 81 c8 c4 e7 4a 39 9d b2 e3 4c 41 25 54 38 d4 ca 91 cf 91 53 42 fb 08 f6 6c b9 3d 9b 73 46 " & _
"61 1b f0 64 30 ff b3 1f 81 e0 c5 2d 45 68 05 19 96 9e 8a a6 33 20 14 3d c8 45 3d aa 40 f5 2f 58 7d 1d 5d c8 7a 17 e6 50 08 3a 7a fd eb 53 49 c0 2e 07 e2 80 48 81 fd ec 68 bf 50 02 0a a0 17 81 " & _
"4c ea 09 b5 4e bb dc e7 4e 95 b2 6f e0 20 52 05 41 01 72 02 94 be fb fd ef 80 d7 89 50 bc e7 9a 84 6c e0 23 81 4f bc e2 17 ff f7 49 6f e4 f1 90 8f bc e4 27 4f f9 ca 5b fe f2 98 ff 47 40 00 3b"
                                                                                                                                                                                                                                                                   

'Call the function to convert hex for a binary write to clients browser
Call displayImage(strHexImage)


'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>