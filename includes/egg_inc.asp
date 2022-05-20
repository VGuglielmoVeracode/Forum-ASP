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









'Dimension variables
Dim sarryHex		'Holds each hex value in the image
Dim lngHexArrayPosition	'Holds the array position
Dim strHex		'Holds the hex stuff

Response.ContentType = "image/gif"

strHex = "" & _
"47 49 46 38 39 61 64 00 55 00 c4 00 00 15 16 18 89 89 86 54 52 55 c7 c0 b0 3c 3a 39 88 98 da 66 66 66 61 71 a9 27 2a 36 77 88 be f7 f3 ef 4b 41 44 ae c1 f4 7d 7a 7b cb e1 ff 58 67 95 59 5c 5e 74 74 7a 28 28 29 78 87 a9 8f 9d c8 68 61 5f 3f 3f 45 6a 7d ae 99 99 99 d9 ef ff 21 21 21 33 33 33 97 ab e4 4d 4a 4c 6b 75 86 d5 db e4 21 f9 04 00 07 00 ff 00 2c 00 00 00 00 64 00 55 00 00 05 ff 60 e6 8c 64 49 66 22 aa ae 6c eb be 62 71 51 14 43 32 0c c7 15 45 e2 4f c0 a0 10 e8 2b 1a 2f c8 83 f2 f0 30 39 1d aa a7 14 66 a2 52 12 14 8e ed 96 e3 f5 12 c8 61 10 49 1e 22 1f e8 74 53 ca 9e c2 56 a7 d6 cd 97 dd 8e 70 3a 9e 2f 2c c4 02 97 4b 64 4c 6a 0f 10 16 6d 6e 6f 2f 55 72 77 05 41 35 38 78 3b 5f 7c 62 7f 07 64 1e 9b 9b 10 9e 9f 16 1a 88 71 8a 8a 4e 8d 38 14 0f 9c 18 34 ae 09 42 07 9c 9c 6a 10 02 b7 9f b9 1d 16 bc 1b 12 00 00 88 a5 c3 28 a3 39 34 13 6a 1e 42 b3 9c 86 bc 16 08 d2 d3 08 1a " & _
"d6 d7 d6 d2 12 1a c0 dd 6d c4 a5 25 92 e3 a9 14 41 84 e8 6a 02 d1 d3 12 ee d8 f0 dd f2 f3 c1 50 e0 f7 22 c7 7b 96 13 17 e9 80 00 d1 74 a0 46 d0 9d c1 6d f1 e8 c9 c3 d7 e8 0e 0e 71 e6 06 a5 53 03 88 8c 45 26 86 08 6a d4 88 4d 21 3d 86 22 b8 90 93 64 6e 22 ba 4c 64 8c ff ec 79 c0 4e e3 86 0d d0 a0 21 38 c8 cd 23 bd 11 54 a4 8c d4 c1 01 59 32 34 80 08 01 54 59 e4 40 c6 8d d2 5e f2 ea 50 41 c0 82 99 06 6b da 94 47 2a a4 b8 91 92 78 fa ec 07 34 20 c0 25 58 5c 5d 31 da 72 e3 cb b3 04 2c 74 58 6b e1 2c 42 6b 53 81 59 bd 8a 75 5c 4f 57 66 be 02 b4 78 21 2c de a3 48 11 10 20 70 b6 70 4c b7 1d a7 3a ac 8b 55 c7 d6 4b 17 be a6 34 52 c0 f1 04 c0 81 11 58 20 bc 41 5a 5a 68 83 07 fb ba 16 17 00 e3 c6 8e cd 4d 60 d2 d7 48 90 22 93 29 57 ee 69 a0 6c e0 97 9c 95 de da c5 59 f4 b6 d2 c0 4e db a5 01 cb 9f 1a 95 7d 88 27 d0 e3 85 47 4f 0f 03 11 9c 0d 3c 78 b3 67 01 bc 71 a3 dd 40 ba b4 f0 9e " & _
"60 4c 12 25 42 c3 0b 4f 4a 3c cc 09 a0 56 78 63 da 0d 04 66 22 d8 f5 56 83 c1 c1 08 81 9b ae db 73 b5 c9 07 44 15 41 43 1e 3a 0c 77 c1 51 86 b5 d5 19 ff 41 85 b9 a3 16 04 2c 69 63 df 74 bf c1 35 15 56 c8 44 f6 5f 80 58 38 c7 13 39 1c 5c c6 4b 68 a1 1d b6 e0 34 b8 c5 a7 59 03 1f 7c f0 40 52 f0 2d a0 60 7d 71 d9 d5 df 45 41 a5 d1 9a 6c b3 7d 98 d5 04 b5 49 13 d3 90 d0 14 66 e4 59 15 7c 40 01 06 2f 4a b7 c1 02 32 e2 07 cc 4b 52 75 43 da 24 3b 10 c5 17 20 45 34 67 5e 81 3f d6 46 e4 98 45 6a 97 62 75 01 7c 10 81 00 06 78 46 22 7e d7 fc 06 80 54 d8 e4 80 9e 97 5a 66 b2 9c 97 3d 6a 81 87 39 cf 90 39 e4 2e a0 99 69 01 94 32 36 30 c0 9a 0b 9c 49 18 7e 08 b9 63 25 37 75 f2 c4 67 73 7b 64 d2 e7 0e e7 9d 47 81 51 84 0e ca 16 34 6b 8d 3a 22 6e 88 a6 8a 41 04 6f 8a 06 1f 01 91 52 09 cc 35 a3 69 60 e9 a5 7a ec 53 40 56 9d 9e 27 43 a0 6a 85 1a 6c a9 6a dd 62 ec 61 04 40 c9 d6 5a 1b 54 c0 " & _
"2a 89 8d ba f3 52 7d bf cc 3a 21 ff 42 62 51 70 29 6c 17 ec 7a 1a 07 60 04 5a 2a 76 a5 ae 65 ec 2e e6 7e 12 ea 60 a9 26 2b c0 60 e4 ee 06 ab 7d 84 75 57 25 4d 00 dc 55 1e 9f 94 09 f7 29 04 a1 8e 4b ae c0 b7 2c 85 8b 27 c2 26 bb 19 a2 c6 0a 50 c1 c3 1d cc ab 01 7c 34 ca 63 5f 54 00 64 ab 1c 51 bb 9a 40 8e 88 4b 95 2b b2 b9 0d f3 22 80 2e 1d c8 48 e4 02 1d 18 fb 70 53 8d 8e 26 01 9c 16 ce 73 71 b5 d9 72 58 99 1d 1e 53 e0 41 5b c3 96 1b b4 c8 b6 98 fc 09 b9 2a ab c5 2e 94 0d 93 ab a0 35 f0 d5 5a a5 95 ef 00 a3 6f b6 5e f8 c0 b3 03 7f 7a f0 6e 5a 2d 0f 0c cd 6e 0d b7 0c 30 01 61 97 0d 0d 94 21 33 7d 6e c4 f5 46 ed 4b d5 f4 24 a6 b1 c6 48 90 63 8e 07 67 57 57 2c b1 06 63 d7 b4 c3 11 77 e0 09 d9 d8 f1 c2 32 c1 03 2f 00 a9 04 0d be 65 f1 9c 35 df 2d 56 7f 13 20 e3 b5 82 69 85 b6 38 e0 16 0c ee b0 ff 2d 15 74 b0 41 07 06 d8 42 b2 d3 69 e3 52 aa e3 d3 1a 04 d3 b4 dd 59 5b 33 00 40 " & _
"b4 92 ad 9f 14 34 60 c0 3a 0a 26 e8 39 db 31 9d 7c 78 2e 45 6f c0 e6 e1 07 27 6e f6 b9 23 72 16 d5 c4 88 d5 3e a7 a4 94 1b 02 41 00 01 b4 c2 d3 aa bc e5 f6 6a 6f c3 bf 19 3a f3 c6 1b f0 3b bc b9 34 6f 2c c2 4c c7 ec d6 5b ee 10 26 67 95 13 c3 a3 a2 c9 ce 94 15 da 91 68 e9 9c ab 4e b7 9b b5 40 a0 02 ea 6b 4a b2 90 67 3c e6 1d 4e 46 0b 70 ca f8 de 61 0d c8 9d e5 76 54 83 c7 4b 32 83 a2 01 1e e9 7f 85 41 db b8 3c 91 40 a7 a0 ed 65 0c b4 c5 f2 08 05 25 f1 4d 6f 66 f6 ab 5d 4d 6e 86 90 e9 70 50 30 20 34 52 0e 71 b3 14 08 a8 ef 80 10 23 8c e1 10 98 42 12 ae e3 2c 2c b3 80 ec 62 35 37 39 4d 8a 72 34 04 60 3b 0e 62 41 00 8e 4f 80 67 0a dd cb 0c 00 b3 c1 0c f1 61 c8 3b 5a c4 0a e3 38 ff 9a 40 4d 6a 75 ab 20 15 67 66 a4 82 50 31 80 a2 f9 1f 94 42 58 9d 93 a9 4f 7d 4e 71 1c 53 ee 78 34 e3 45 ac 17 b2 bb 20 d4 ea 95 18 9b d1 f0 3e 6d a4 c6 1a ab 48 c7 33 8d 4f 29 86 e3 23 c0 1a 65 01 " & _
"04 26 f0 78 9e e0 5c 13 11 83 48 3a 79 84 82 07 79 13 54 16 79 10 38 b6 ea 83 04 60 d3 0f 8d e5 b8 54 a2 b0 61 b6 70 a1 05 a3 12 c8 a9 29 c4 3e 6a 74 50 91 48 59 4a 3a 92 c8 8a 31 82 40 04 22 d0 94 dd a4 ca 78 88 03 98 2f b4 23 ad 35 d6 ea 96 f7 7a 63 a8 a6 48 ca 43 a1 b2 37 3a 4c 56 05 7c 57 cc 3c b6 d0 70 05 34 5c 46 a4 15 bb f9 e1 4b 86 4e bc 5e 29 cb 04 3b 52 9e ce 74 8c 7c 95 15 d9 25 cc 0a 1c ad 5d 61 2b 57 26 61 b4 c9 ea d1 c9 89 18 ec e5 3b 37 b3 16 f9 c5 ae 7e 0f e3 4e 36 1b f9 aa 05 20 f0 65 d8 51 96 e3 58 36 30 b3 19 f4 a0 b3 84 87 f5 14 42 ce b3 b0 ff c5 a1 15 80 9d 39 25 c0 94 0a 10 40 03 4e 02 e6 99 98 42 4c 87 21 ad 85 a9 1a e1 53 dc 49 3b 4a 61 f0 96 d2 92 5e a9 b6 c9 36 32 c6 cc 61 26 ed a8 a3 e2 e8 b9 6d 42 b4 a0 e1 0b 9a f1 66 4a 45 1a de d4 62 53 a3 9e 91 c8 65 80 06 ac 25 37 20 25 0c 50 61 75 c5 56 b5 ca a1 be c3 63 ca 5a 96 32 b7 bd 4f 99 6b 8c d3 46 " & _
"0d 69 af 23 0d eb 61 b7 88 cf 93 5e 26 c4 87 11 66 8e 5e f5 2a 9b 88 09 57 b3 96 75 70 a5 e3 65 c5 6c 67 53 77 38 8e 52 f0 19 92 ea d2 e6 ae 04 9a d4 5d 70 93 60 36 bd ba cd 06 1c 35 6d 65 15 99 c3 16 20 58 74 a2 8d 3b 17 9b 18 94 7e 01 b9 cf a8 c5 02 a9 13 d9 02 aa da 80 06 04 a0 01 2f 21 5c 6c e1 96 d7 98 2c a0 b5 5c 84 98 b9 e0 5a 56 44 89 b3 03 4b a4 9f 2f ac a4 cd 08 00 77 42 1a 20 40 48 b9 03 b9 21 11 e0 87 45 7b d0 1d 07 10 80 97 30 05 b8 ff 12 b0 ab ab be 0a a5 ca 5e f6 8b de f4 ed c9 04 d0 44 0a 26 24 95 7b 35 1d 6e 24 e0 d0 d1 4a c0 b9 cf bd 63 6a 7b 58 cf bb 3a 05 72 ad 65 57 5e 29 ea 58 de b2 0c a2 a9 f2 ed 01 95 08 39 f3 76 24 bb 4d b1 27 c0 0c a7 44 a6 cc 11 26 9b 39 ad 05 22 a0 be 08 90 90 37 41 3b 61 a3 98 f2 b5 16 26 8b 6d 4c fb e1 cb 60 c6 32 20 f6 36 c0 87 6b 50 42 00 d0 de d5 f2 11 9e 6b c9 a9 57 1b 30 4c f9 2e b8 44 9b 11 e6 58 db c2 5d 95 01 15 ae 40 " & _
"fd 63 04 ed d9 ca a5 29 8b ac 4d d4 a8 35 12 3a e4 d4 56 23 b9 b0 82 a1 df 94 42 63 0a d7 f8 70 c2 b2 80 0f e1 36 d4 37 2d 80 79 47 8d e9 2d 8a ec b7 c0 05 4f c9 1a 78 17 0c db c9 48 d3 8a 6a 98 70 a6 b0 7c 1b e8 b4 a1 7a 38 2d 14 55 e1 b9 1c b7 b6 82 92 b9 44 24 9b 91 92 7f 01 0f 69 81 86 5d e8 82 db 2e 7c 07 e7 d6 c2 59 be 77 5c ff 07 e8 36 e3 66 b3 a5 ae 6c 64 5e dc 9f cb 1c ba 23 d2 ef d3 81 8c 89 a9 4c 15 13 39 47 c0 b5 8f ae b0 8d d5 75 e8 0f 27 ab 65 96 44 5f ca 44 1d 3e 2f ff 75 1d 12 9b a5 8c 0d 37 b2 96 2d 2c d1 eb aa b2 a3 23 80 81 00 18 20 d5 72 e6 63 e2 28 bd b6 2f 5f 52 d6 48 f5 33 f1 14 27 b0 88 d5 70 5a 3c 7c 10 c0 ce 75 32 cd 16 0c 1a ad 1d f6 a9 69 7c ec 1a 9b da 9e 62 55 5c d2 1c 0a 5d 3d 0b ae d7 e5 8a a9 38 c9 5b c3 32 8e af da 3e d4 f3 df be bd 94 70 c3 99 7b b8 65 2d 85 7d 98 6c 84 1d 2a c0 7b b4 67 38 57 e7 ed 8a a6 4b 99 cb 44 1b ac dc 6c b4 79 1f " & _
"0b 9c d8 21 e0 3a 84 ed da d7 56 79 dc c9 3e b6 b2 9b 8d 28 1f f2 76 b7 ed 6b 9a c1 cb 05 bd dc b4 e5 bd ad 2b a0 ea 48 b5 58 74 ed c6 ca c3 1c 76 b8 1d 7d c7 53 a7 ce e0 46 86 12 10 15 0c 6b 05 87 d1 16 58 ae 63 c3 ff 44 0a 13 69 3c 5c 85 31 37 57 ea 42 c6 bc 55 0d 13 e0 be db b9 a3 11 18 f0 d4 86 0f a6 40 6d 18 90 51 78 40 77 17 ec 74 c5 6c 4a f0 92 25 24 3a 23 ac 54 06 47 e6 b2 8c 37 80 d6 62 dd ee 34 06 78 9c e3 fc bb e8 c5 14 81 62 2f 66 d7 a1 1b af 5d 88 1d 9e a8 9a 99 01 6d 81 c7 02 8e 37 9f 6c 11 e1 2d 30 d0 5a ca 1f 9b 7b af 3d b5 95 03 b0 f7 1f 46 6f 69 25 ed e6 27 9c 95 c0 72 53 f8 6d 29 bb 85 5d c9 99 16 c5 a3 0e d2 ef 3b 98 ba 08 65 73 8f 67 5e dc 7b df f9 95 01 46 69 cf a9 be 9b 23 be e3 cb 6a 8c e9 26 0b a0 7e f5 f3 45 ba 8e 4d c2 76 ab b2 65 91 26 6b cb 5e db bd 62 6b bd e3 34 d6 bd b2 69 8b e8 b0 c3 92 ec 9e 00 e3 9e 43 6c 08 18 96 76 62 66 f3 e1 ef 54 47 " & _
"61 8c bb 4f f6 10 18 40 b1 31 cf f9 06 cc 9f 7b 8d 0e ab f0 b1 e3 e5 21 37 85 58 5e 84 38 48 33 2a ff 4c d3 14 f0 81 78 dc 21 42 af f7 73 db a6 3a 08 f3 76 e9 67 2c 1f 20 7f 98 47 79 ae 55 6c 38 d7 79 c3 34 73 3c 16 1a 45 27 2f db f0 24 47 16 65 52 f6 61 9e 50 64 f3 d3 69 e9 d3 7c 47 a3 7e 0b d6 40 b9 a0 00 13 58 6c 18 60 7d f6 f7 5a 18 40 5d 11 e0 71 61 45 61 bc d5 7f 0d 73 6c f4 f6 61 65 54 41 22 d4 5d a5 83 1f 0d f2 4e c2 b4 7e 0e 64 47 67 85 0b eb 10 83 32 28 7f 35 78 7f f6 a7 00 75 b7 83 af 75 47 8e e6 3b b8 46 22 ad 53 55 bf 23 6d 9c 01 01 a8 e6 14 f5 04 4f be e1 0e e2 a4 7e a8 07 7d 0e d4 47 d8 a1 3a 0a 70 87 03 90 87 7a 88 79 79 c8 83 3a 17 6e a0 00 2d 8b 03 4b 08 03 2d a8 c3 79 63 d5 57 66 12 82 4d 98 3a 62 45 88 e9 f3 3b a2 b3 16 77 78 87 2d 32 81 79 d8 71 f4 e7 6f ba a7 7e b5 a6 2c e3 87 28 29 f2 39 83 01 44 19 17 35 f7 61 47 35 d6 6e 24 f4 09 ff ca 16 2f 4d 53 " & _
"89 b2 b8 87 f4 57 8b d8 c7 73 65 d8 7b 9f f3 7f 9b f6 4b cd 32 70 92 15 3b 0b 58 6e ea 27 49 c6 78 69 65 b3 1b 10 20 8b b3 68 85 58 88 79 78 67 77 a7 b7 6c af 66 71 63 06 42 88 46 7a 7d d7 86 68 f3 43 90 e6 8d 91 96 6f ad 28 46 cb c3 26 cc 58 89 2d e2 8c 9b c8 83 5c 98 79 44 17 3d dc 06 54 10 d0 4a 5f f6 30 c7 56 4c 29 03 29 04 54 8c 0c b8 57 91 06 83 e2 88 65 fc 78 8e b2 d8 22 f3 57 83 b5 88 75 1e 67 61 ea 93 1d 7d 16 49 a7 67 49 90 d6 5b bd 11 3a e0 f8 62 22 87 47 c8 c3 80 66 f7 09 04 c9 8c 98 78 85 d4 87 79 39 e7 6f 72 66 42 21 34 36 86 a7 60 f6 54 50 8f e2 45 cd 97 3a 0c 89 8b ed d6 6e 03 c7 38 1e f9 91 20 69 85 57 07 8d 5a c7 3d 23 57 1d a1 21 1d a2 f1 6b f5 03 3b e2 74 6c e4 16 8e f2 b5 81 27 b3 8a 28 b3 3a 47 a3 93 e7 98 8e 07 89 90 a7 ae 55 65 f8 17 7d f9 b4 6c 05 43 18 d0 90 49 c1 e2 85 be 53 8c 91 a8 6a 91 04 5d 24 83 65 4f 47 95 3a 39 83 d6 67 8b e1 26 7c 12 99 " & _
"5a b7 50 86 a9 27 62 0c 58 61 72 96 0b 3d c7 90 92 68 47 7a f6 86 f3 e5 82 6e 09 92 b3 88 90 08 69 8b 38 f7 8d ed 36 62 19 e8 98 aa 56 6e 90 e6 8a ab d6 34 8f 78 38 87 f9 96 7a 98 83 00 47 7f 7b 97 6a e4 e6 97 90 36 70 04 a7 91 15 16 70 e2 d8 73 4e 48 87 de 38 95 9b 59 95 2d 52 85 14 98 90 59 b9 90 7c 19 72 1b 48 99 76 54 55 7d 89 0b 21 f7 73 af 49 9a 0c 24 7b 21 00 00 3b"


'Place the Hex into an array
sarryHex = split(strHex, " ")

'Loop through the array to convert to binary
For lngHexArrayPosition = 0 to CLng(UBound(sarryHex))

	'Write the binary output
	Response.BinaryWrite ChrB(CLng("&H" & sarryHex(lngHexArrayPosition)))
Next

Response.Flush
Response.End


%>