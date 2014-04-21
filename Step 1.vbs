'============================================
'Code HTML+CSS from Image Files
'============================================

'::Settings
strTitle      = "Image Coder v1.7.4"
strFileName   = Wscript.ScriptName
strFilePath   = Replace(WScript.ScriptFullName,strFileName,"")
strImagesPath = "images/"
strStylePath  = "styles/style.css"
strIISPath    = "C:\Windows\System32\inetsrv\config\applicationHost.config"
strRegPath    = "HKEY_CURRENT_USER\Software\Adobe\Common\11.5\Sites\" 'CS5.5
strRegPath2   = "HKEY_CURRENT_USER\Software\Adobe\Common\12.0\Sites\" 'CS6
strRegPath3   = "HKEY_CURRENT_USER\Software\Adobe\Common\13.0\Sites\" 'CC
strRegSite    = "-Site0"
strRegName    = "PCM"


'::Parse Files
intCount  = 0
intAnswer = Msgbox("Parse Files?",vbYesNo,strTitle)
If intAnswer = vbYes Then

	'::Get Navigation Order
	strNavOrder = InputBox("List Pages in Order:","Navigation","home, services, about, contact")
	strNavOrder = LCase(strNavOrder)
	arrNavOrder = Split(strNavOrder,", ")

	strNavLabel = InputBox("List Page Labels in Order:","Navigation",PCase(strNavOrder))
	arrNavLabel = Split(strNavLabel,", ")


	'::Load Stylesheet
	strCSS = ReadFile(strFilePath & strStylePath)


	'::Find Index/Default File
	If FileExists(strFilePath & "default.asp") Then
		strFileIndex = "default.asp"
	ElseIf FileExists(strFilePath & "index.asp") Then
		strFileIndex = "index.asp"
	Else
		strFileIndex = "index.php"
	End If
	strFileIndexExt  = LCase(Mid(strFileIndex,InStrRev(strFileIndex,".")+1))
	strHTML = ReadFile(strFilePath & strFileIndex)


	'::Get Images
	Set objFSO    = CreateObject("Scripting.FileSystemObject")
	Set objFolder = objFSO.GetFolder(strFilePath & strImagesPath)
	Set objFiles  = objFolder.Files

	Set objDicCSS  = CreateObject("Scripting.Dictionary")
	Set objDicHead = CreateObject("Scripting.Dictionary")
	Set objDicFoot = CreateObject("Scripting.Dictionary")

	For Each objFile In objFiles
		strFileFull   = objFile.Name
		strFileExt    = LCase(Mid(strFileFull,InStrRev(strFileFull,".")+1))
		strFileName   = Mid(strFileFull,1,InStrRev(strFileFull,".")-1)
		strFileSize   = objFile.Size
		strFileDateC  = objFile.DateCreated
		strFileDateLA = objFile.DateLastAccessed
		strFileDateLM = objFile.DateLastModified
		strFileAll    = objFile.Path
		strFilePre    = Mid(strFileFull,1,InStr(strFileFull,"_"))

		'Layout Files
		If strFilePre = "Layout_" Then
			Select Case strFileFull
				Case "Layout_Header.jpg"
					strCSS = UpdateCSS(strCSS,strFileAll,"#header","height")

				Case "Layout_Footer.jpg"
					strCSS = UpdateCSS(strCSS,strFileAll,"#footer","height")

				Case "Layout_Body.jpg"
					strCSS = UpdateCSS(strCSS,strFileAll,"#body_container","min-height")
					strCSS = Replace(strCSS,_
							 "/*background: url(../images/Layout_Body.jpg) no-repeat center top;*/",_
							 "background: url(../images/Layout_Body.jpg) no-repeat center top;")

				Case "Layout_BG.jpg"
					strCSS = Replace(strCSS,_
							 "/*background: url(../images/Layout_BG.jpg) repeat-x;*/",_
							 "background: url(../images/Layout_BG.jpg) repeat-x;")

				Case "Layout_BodyBG.jpg"
					strCSS = Replace(strCSS,_
							 "/*background: url(../images/Layout_BodyBG.jpg) repeat-y center top;*/",_
							 "background: url(../images/Layout_BodyBG.jpg) repeat-y center top;")

				Case "Layout_FooterBG.jpg"
					strCSS = Replace(strCSS,_
							 "/*background: url(../images/Layout_FooterBG.jpg) repeat-x center top;*/",_
							 "background: url(../images/Layout_FooterBG.jpg) repeat-x center top;")

			End Select

			intCount = intCount + 1

		'Navigation Files
		ElseIf strFilePre = "Nav_" Then
			strFilePost = Mid(strFileName,InStr(strFileName,"_")+1)

			'CSS (Base)
			strCSS = UpdateCSS(strCSS,strFileAll,"#nav a","height")
			strCSS = UpdateCSS(strCSS,strFileAll,"#nav a.selected","background-position")

			'CSS (Header)
			strTmpCSS  = "#nav a." & LCase(strFilePost) & " {" & vbCrlf & vbTab & _
						 "background-image: url(../" & strImagesPath & strFileFull & ");" & vbCrlf & vbTab & _
						 "width: " & ImageSize(strFileAll)(0) & "px;" & vbCrlf & _
						 "}" & vbCrlf

			'HTML (Header)
			If strFilePost = "Home" Then
				strURLPost = Mid(strFileIndex,1,InStrRev(strFileIndex,".")-1)
			Else
				strURLPost = strFilePost
			End If

			'HTML (Nav Label)
			strLabel = strFilePost
			For i = 0 To UBound(arrNavOrder)
				If LCase(strFilePost) = arrNavOrder(i) Then
					strLabel = arrNavLabel(i)
					Exit For
				End If
			Next

			strTmpHead = String(4,vbTab) & "<li><a href=""" & LCase(strURLPost) & "." & strFileIndexExt & """ class=""" & LCase(strFilePost)

			If strFileIndexExt = "asp" Then
				strTmpHead = strTmpHead & "<" & "%=NavSel(strCurFile,""" & LCase(strURLPost) & """)%" & ">"
			Else
				strTmpHead = strTmpHead & "<?php echo navSel($strPage,'" & LCase(strURLPost) & "'); ?>"
			End If

			strTmpHead = strTmpHead & """>" & strLabel & "</a></li>" & vbCrlf

			'HTML (Footer)
			strTmpFoot = String(4,vbTab) & "<li><a href=""" & LCase(strURLPost) & "." & strFileIndexExt & """>" & strLabel & "</a>|</li>" & vbCrlf

			strNavCSS  = strNavCSS & strTmpCSS
			strNavHead = strNavHead & strTmpHead
			strNavFoot = strNavFoot & strTmpFoot

			objDicCSS.Add LCase(strFilePost),strTmpCSS
			objDicHead.Add LCase(strFilePost),strTmpHead
			objDicFoot.Add LCase(strFilePost),strTmpFoot

			intCount = intCount + 1
		End If
	Next

	Set objFiles  = Nothing
	Set objFolder = Nothing
	Set objFSO    = Nothing


	'::Insert CSS Nav
	If strTmpCSS <> "" Then
		strTmp = strCSS
		strTmp = Mid(strTmp,InStr(strTmp,"#nav ul li {"))
		strTmp = Mid(strTmp,1,InStr(strTmp,"/* Navigation (Footer) */") + 25)

		strAtt = strTmp
		strAtt = Mid(strAtt,InStr(strAtt,"}")+1)
		strAtt = Mid(strAtt,1,InStr(strAtt,"/*")-1)

		strNew = vbCrlf & vbCrlf

		For i = 0 To UBound(arrNavOrder)
			If objDicCSS.Exists(arrNavOrder(i)) Then
				strNew = strNew & objDicCSS(arrNavOrder(i))
				objDicCSS.Remove(arrNavOrder(i))
			End If
		Next

		arrDicCSS = objDicCSS.Items
		For i = 0 To objDicCSS.Count-1
			strNew = strNew & arrDicCSS(i)
		Next

		strNew = strNew & vbCrlf & vbCrlf

		strNew = Replace(strTmp,strAtt,strNew)
		strCSS = Replace(strCSS,strTmp,strNew)
	End If


	'::Insert Header Nav
	If strNavHead <> "" AND InStr(strHTML,"<div id=""nav"">") Then
		strTmp = strHTML
		strTmp = Mid(strTmp,InStr(strTmp,"<div id=""nav"">"))
		strTmp = Mid(strTmp,1,InStr(strTmp,"</div>") + 6)

		strAtt = strTmp
		strAtt = Mid(strAtt,InStr(strAtt,"<ul>")+4)
		strAtt = Mid(strAtt,1,InStr(strAtt,"</ul>")-1)

		strNew = vbCrlf

		For i = 0 To UBound(arrNavOrder)
			If objDicHead.Exists(arrNavOrder(i)) Then
				strNew = strNew & objDicHead(arrNavOrder(i))
				objDicHead.Remove(arrNavOrder(i))
			End If
		Next

		arrDicHead = objDicHead.Items
		For i = 0 To objDicHead.Count-1
			strNew = strNew & arrDicHead(i)
		Next

		strNew = strNew & String(3,vbTab)

		strNew = Replace(strTmp,strAtt,strNew)
		strHTML = Replace(strHTML,strTmp,strNew)
	End If


	'::Insert Footer Nav
	If strNavFoot <> "" AND InStr(strHTML,"<div id=""navf"">") Then
		strTmp = strHTML
		strTmp = Mid(strTmp,InStr(strTmp,"<div id=""navf"">"))
		strTmp = Mid(strTmp,1,InStr(strTmp,"</div>") + 6)

		strAtt = strTmp
		strAtt = Mid(strAtt,InStr(strAtt,"<ul>")+4)
		strAtt = Mid(strAtt,1,InStr(strAtt,"</ul>")-1)

		strNew = vbCrlf

		For i = 0 To UBound(arrNavOrder)
			If objDicFoot.Exists(arrNavOrder(i)) Then
				strNew = strNew & objDicFoot(arrNavOrder(i))
				objDicFoot.Remove(arrNavOrder(i))
			End If
		Next

		arrDicFoot = objDicFoot.Items
		For i = 0 To objDicFoot.Count-1
			strNew = strNew & arrDicFoot(i)
		Next

		strNew = strNew & String(3,vbTab)

		If strFileIndexExt = "asp" Then
			strNew = strNew & String(4,vbTab) & "<li><a href=""sitemap.asp"">Site Map</a></li>"
		Else
			strNew = Left(strNew,Len(strNew)-11) & "</li>"
		End If
		strNew = strNew & vbCrlf & String(3,vbTab)

		strNew = Replace(strTmp,strAtt,strNew)
		strHTML = Replace(strHTML,strTmp,strNew)
	End If

	Set objDicCSS  = Nothing
	Set objDicHead = Nothing
	Set objDicFoot = Nothing


	'::Set CSS Values
	strPreValue  = ""
	intAnswer    = Msgbox("Assign CSS Values?",vbYesNo,strTitle)
	If intAnswer = vbYes Then
		Call SetCSS("Background Color (body)","body","background-color","#FFF")

		Call SetCSS("Heading 1 Color (h1)","h1|.Heading1|.Highlight3|.Btn1|.Background1|.Callout1","color","#CC0000")
		Call SetCSS("Link Normal Color (a)","a|#faq dt","color",strPreValue)
		Call SetCSS("Form Input Border Color","input:focus, select:focus, textarea:focus","border-color",strPreValue)
		Call SetCSS("Form Input Glow Color","input:focus, select:focus, textarea:focus","box-shadow",strPreValue)
		Call SetCSS("Heading 1 Size (h1)","h1","font-size","30px")

		Call SetCSS("Heading 2 Color (h2)","h2|.Heading2|.Highlight4|.Btn2|.Background2|.Callout2","color","#CC0000")
		Call SetCSS("Link Rollover Color (a:hover)","a:hover|a:hover .ImageLink|#navf ul li a:hover|#copyright a:hover|#designed a:hover|#faq dt:hover","color",strPreValue)
		Call SetCSS("Heading 2 Size (h2)","h2","font-size","25px")

		Call SetCSS("Heading 3 Color (h3)","h3|.Heading3|.Highlight5|.Btn3|.Background3|.Callout3","color","#CC0000")
		Call SetCSS("Heading 3 Size (h3)","h3","font-size","20px")

		Call SetCSS("Heading 4 Color (h4)","h4|.Heading4|.Highlight6|.Btn4|.Background4|.Callout4","color","#CC0000")
		Call SetCSS("Heading 4 Size (h4)","h4","font-size","18px")

		Call SetCSS("Heading 5 Color (h5)","h5|.Heading5|.Highlight7|.Btn5|.Background5|.Callout5","color","#CC0000")
		Call SetCSS("Heading 5 Size (h5)","h5","font-size","16px")

		Call SetCSS("Navigation Overlay Background Color","NAV-BG","","#252525")
		Call SetCSS("Navigation Overlay Foreground Color","NAV-FG","","#FFCC00")

		Call SetCSS("Body Text Color (body, input)","body|input, select, textarea","color","#000")
		Call SetCSS("Body Text Size (body, input)","body|input, select, textarea","font-size","14px")

		Call SetCSS("Footer Navigation Color (#navf)","#navf|#navf ul li a","color","#FFF")
		Call SetCSS("Footer Copyright Color (#copyright, #designed)","#copyright|#copyright a|#designed|#designed a","color",strPreValue)
		Call SetCSS("Footer Navigation Size (#navf)","#navf","font-size","14px")
		Call SetCSS("Footer Copyright Size (#copyright, #designed)","#copyright|#designed","font-size",strPreValue)
	End If


	'::Remove Trailing CRLFs
	intCrlf = InStrRev(strCSS,"}")
	strCSS  = Left(strCSS,intCrlf)

	intCrlf = InStrRev(strHTML,">")
	strHTML = Left(strHTML,intCrlf)


	'::Save Files
	intAnswer = Msgbox("Update " & strStylePath & " file?",vbYesNo,strTitle)
	If intAnswer = vbYes Then
		Call WriteFile(strFilePath & strStylePath,strCSS,"Write")
	End If

	intAnswer = Msgbox("Update " & strFileIndex & " file?",vbYesNo,strTitle)
	If intAnswer = vbYes Then
		Call WriteFile(strFilePath & strFileIndex,strHTML,"Write")
	End If

End If


intAnswer = Msgbox("Update IIS Config?",vbYesNo,strTitle)
If intAnswer = vbYes Then
	strIIS = ReadFile(strIISPath)

	Set objRegExp = New Regexp
	objRegExp.IgnoreCase = True
	objRegExp.Global     = True
	objRegExp.Pattern    = "<site.*>\s*<application.*>\s*<virtualDirectory.*physicalPath=""(.*?)""\s/>"
	strIIS               = objRegExp.Replace(strIIS,"<site name=""Default Web Site"" id=""1"">" & vbCrlf & Space(16) & _
						   "<application path=""/"">" & vbCrlf & Space(20) & _
						   "<virtualDirectory path=""/"" physicalPath=""" & Left(strFilePath,Len(strFilePath)-1) & """ />")
	Set objRegExp        = Nothing

	Call WriteFile(strIISPath,strIIS,"Write")
End If

intAnswer = Msgbox("Update Dreamweaver Config?",vbYesNo,strTitle)
If intAnswer = vbYes Then
	Dim WSHShell
	Set WSHShell = WScript.CreateObject("WScript.Shell")

	strFilePathReg  = strFilePath
	strFilePathReg  = Replace(strFilePathReg,"\","\")
	strImagePathReg = strFilePathReg & Replace(strImagesPath,"/","\")

	WSHShell.RegWrite strRegPath & strRegSite & "\" & "Local Directory", strFilePathReg
	WSHShell.RegWrite strRegPath & strRegSite & "\" & "Image Directory", strImagePathReg
	WSHShell.RegWrite strRegPath & "-Summary\" & "Current Site", strRegName

	WSHShell.RegWrite strRegPath2 & strRegSite & "\" & "Local Directory", strFilePathReg
	WSHShell.RegWrite strRegPath2 & strRegSite & "\" & "Image Directory", strImagePathReg
	WSHShell.RegWrite strRegPath2 & "-Summary\" & "Current Site", strRegName

	WSHShell.RegWrite strRegPath3 & strRegSite & "\" & "Local Directory", strFilePathReg
	WSHShell.RegWrite strRegPath3 & strRegSite & "\" & "Image Directory", strImagePathReg
	WSHShell.RegWrite strRegPath3 & "-Summary\" & "Current Site", strRegName

	Set WSHShell = Nothing
End If


'::Done
MsgBox "Finished processing " & intCount & " images.", vbInformation



'============================================
'Functions / Sub-Procedures
'============================================

'::Read File
Function ReadFile(filepath)
	Set fso = Createobject("Scripting.FileSystemObject")
	Set tmpRead = fso.OpenTextFile(filepath)
	While Not tmpRead.AtEndOfStream
		strRead = tmpRead.ReadAll
	Wend
	Set tmpRead = Nothing
	Set fso = Nothing
	ReadFile = strRead
End Function


'::Write File
'ex.WriteFile("c:\file.ext","text","Append")
Sub WriteFile(filepath,text,mode)
	Select Case mode
		Case "Write"
			fsoMode = 2
		Case "Append"
			fsoMode = 8
	End Select
	Set fso = Createobject("Scripting.FileSystemObject")
	Set fsoWrite = fso.OpenTextFile(filepath, fsoMode, True)
	fsoWrite.WriteLine(text)
	fsoWrite.Close
	Set fso = Nothing
End Sub


'::Check if File Exists
Function FileExists(filepath)
	Set fso = Createobject("Scripting.FileSystemObject")
	FileExists = fso.FileExists(filepath)
	Set fso = Nothing
End Function


'::RegEx Function
Function RegEx(strInput,strExp,strMode)
	Set objRegExp = New Regexp
	objRegExp.IgnoreCase = True
	objRegExp.Global     = True
	objRegExp.MultiLine  = False
	objRegExp.Pattern    = strExp

	objOutput = ""
	strOutput = ""
	Select Case LCase(strMode)
		Case "matches" 'array
			For Each objMatch In objRegExp.Execute(strInput)
				strOutput = strOutput & objMatch.Value & "|"
			Next

			If Right(strOutput,1) = "|" Then
				strOutput = Left(strOutput,Len(strOutput)-1)
			End If

			objOutput = Split(strOutput,"|")

		Case "groups"  'array:  (.*?)
			For Each objMatch In objRegExp.Execute(strInput)
				For intMatch = 0 To objMatch.SubMatches.Count-1
					strOutput = strOutput & objMatch.SubMatches(intMatch) & "|"
				Next
				If Right(strOutput,1) = "|" Then
					strOutput = Left(strOutput,Len(strOutput)-1)
				End If
				strOutput = strOutput & "[]"
			Next

			If Right(strOutput,2) = "[]" Then
				strOutput = Left(strOutput,Len(strOutput)-2)
			End If

			objOutput = Split(strOutput,"[]")

		Case "test"    'boolean
			objOutput = objRegExp.Test(strInput)

		Case "replace" 'string: <(.|\n)+?>
			objOutput = objRegExp.Replace(strInput,"")

	End Select

	Set objRegExp = Nothing
	RegEx = objOutput
End Function


'::Image Size Function (LoadPicture)
'ex. intImgWidth = ImageSize("images/photo.jpg")(0)
Function ImageSize(strPath)
	arrImg     = Array("","")
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	'If NOT InStr(strPath,":") Then strPath = Server.MapPath("/" & strPath)
	If NOT objFSO.FileExists(strPath) Then Exit Function

	blnImg = ImgInfo(strPath, intImgWidth, intImgHeight, intImgColors, strImgType)
	arrImg(0) = intImgWidth
	arrImg(1) = intImgHeight

	Set objFSO = Nothing
	ImageSize = arrImg
End Function


'::Update CSS
Function UpdateCSS(strCSS,strStyleVal,strStyleID,strStyleAttr)
	strTmp = strCSS
	strTmp = Mid(strTmp,InStr(strTmp,vbCrlf & strStyleID & " {"))
	strTmp = Mid(strTmp,1,InStr(strTmp,"}"))

	strAtt = strTmp
	strAtt = Mid(strAtt,InStr(strAtt,vbTab & strStyleAttr & ": "))
	strAtt = Mid(strAtt,1,InStr(strAtt,";"))

	If InStr(LCase(strStyleAttr),"width") Then
		intWH = 0
	Else
		intWH = 1
	End If

	If strStyleAttr = "background-position" Then
		strNew = strStyleAttr & ": 0 -" & Round(ImageSize(strStyleVal)(intWH) / 2) & "px !important;"

	ElseIf strStyleID = "#nav a" Then
		strNew = strStyleAttr & ": " & Round(ImageSize(strStyleVal)(intWH) / 2) & "px;"

	ElseIf strStyleAttr = "box-shadow" Then
		If Left(strStyleVal,1) <> "#" Then strStyleVal = "#" & strStyleVal
		strNew = strStyleAttr & ": 0 1px 2px rgba(0,0,0,0.15) inset, 0 0 3px " & strStyleVal & " !important;"

	ElseIf strStyleAttr = "text-shadow" Then
		If Left(strStyleVal,1) <> "#" Then strStyleVal = "#" & strStyleVal
		strNew = strStyleAttr & ": 0 -1px 0 " & strStyleVal & ";"

	ElseIf strStyleAttr = "border-left" Then
		If Left(strStyleVal,1) <> "#" Then strStyleVal = "#" & strStyleVal
		strNew = strStyleAttr & ": 7px solid " & strStyleVal & ";"

	ElseIf strStyleAttr = "border-top" Then
		If Left(strStyleVal,1) <> "#" Then strStyleVal = "#" & strStyleVal
		strNew = strStyleAttr & ": 1px solid " & strStyleVal & ";"

	ElseIf strStyleAttr = "border-bottom" Then
		If Left(strStyleVal,1) <> "#" Then strStyleVal = "#" & strStyleVal
		strNew = strStyleAttr & ": 1px solid " & strStyleVal & ";"

	ElseIf InStrRev(strStyleVal,"\") > 0 Then
		strNew = strStyleAttr & ": " & ImageSize(strStyleVal)(intWH) & "px;"

	ElseIf InStr(strStyleAttr,"color") Then
		If Left(strStyleVal,1) <> "#" Then strStyleVal = "#" & strStyleVal
		strNew = strStyleAttr & ": " & UCase(strStyleVal) & ";"

	ElseIf IsNumeric(strStyleVal) Then
		strNew = strStyleAttr & ": " & strStyleVal & "px;"

	Else
		strNew = strStyleAttr & ": " & strStyleVal & ";"

	End If

	strNew    = Replace(strTmp,strAtt,vbTab & strNew)
	UpdateCSS = Replace(strCSS,strTmp,strNew)
End Function


'::Set CSS Values
Sub SetCSS(strDesc,strID,strAttr,strValue)
	strNewValue = InputBox(strDesc & ":","CSS " & strAttr,strValue)
	strPreValue = strNewValue

	arrID = Split(strID,"|")
	For Each strID In arrID
		If strID = "a:hover .ImageLink" Then
			strCSS = UpdateCSS(strCSS,strNewValue,strID,"border-color")

		ElseIf Left(strID,4) = ".Btn" Then
			strCSS = UpdateCSS(strCSS,ColorBrightness(strNewValue,0,2),strID,"background-color")
			strCSS = UpdateCSS(strCSS,ColorBrightness(strNewValue,-20,2),strID,"border-color")
			strCSS = UpdateCSS(strCSS,ColorBrightness(strNewValue,-16,2),strID & ":hover","border-color")
			strCSS = UpdateCSS(strCSS,ColorBrightness(strNewValue,6,2),strID & ":hover","background-color")
			'strCSS = UpdateCSS(strCSS,ColorBrightness(strNewValue,98,1),strID,"color")

		ElseIf Left(strID,11) = ".Background" Then
			strCSS = UpdateCSS(strCSS,ColorBrightness(strNewValue,95,1),strID,"background-color")
			strCSS = UpdateCSS(strCSS,ColorBrightness(strNewValue,86,1),strID,"border-color")
			strCSS = UpdateCSS(strCSS,ColorBrightness(strNewValue,-10,2),strID,"color")

		ElseIf Left(strID,8) = ".Callout" Then
			strCSS = UpdateCSS(strCSS,ColorBrightness(strNewValue,95,1),strID,"background-color")
			strCSS = UpdateCSS(strCSS,ColorBrightness(strNewValue,86,1),strID,"border-color")
			strCSS = UpdateCSS(strCSS,ColorBrightness(strNewValue,-10,2),strID,"color")

		ElseIf Left(strID,11) = "#paging_nav" Then
			strCSS = UpdateCSS(strCSS,ColorBrightness(strNewValue,95,1),strID,"background-color")
			strCSS = UpdateCSS(strCSS,ColorBrightness(strNewValue,86,1),strID,"border-color")
			strCSS = UpdateCSS(strCSS,ColorBrightness(strNewValue,-10,2),strID,"color")

		ElseIf Left(strID,6) = "NAV-BG" Then
			strCSS = UpdateCSS(strCSS,strNewValue,"#nav ul ul","background-color")
			strCSS = UpdateCSS(strCSS,ColorBrightness(strNewValue,-10,2),"#nav li ul a","text-shadow")
			strCSS = UpdateCSS(strCSS,ColorBrightness(strNewValue,-10,2),"#nav ul ul","border-left")
			strCSS = UpdateCSS(strCSS,ColorBrightness(strNewValue,5,2),"#nav ul li li","border-top")
			strCSS = UpdateCSS(strCSS,ColorBrightness(strNewValue,-5,2),"#nav ul li li","border-bottom")
			strCSS = UpdateCSS(strCSS,ColorBrightness(strNewValue,8,2),"#nav li ul a:hover","background-color")
			strCSS = UpdateCSS(strCSS,ColorBrightness(strNewValue,10,2),"#nav ul li.parent > a","border-color")
			strCSS = UpdateCSS(strCSS,ColorBrightness(strNewValue,10,2),"#nav ul li.parent > a > div.marker","color")

		ElseIf Left(strID,6) = "NAV-FG" Then
			strCSS = UpdateCSS(strCSS,strNewValue,"#nav li ul a:hover","color")
			strCSS = UpdateCSS(strCSS,strNewValue,"#nav ul li.parent > a:hover","border-color")
			strCSS = UpdateCSS(strCSS,strNewValue,"#nav ul li.parent > a:hover > div.marker","color")

		Else
			strCSS = UpdateCSS(strCSS,strNewValue,strID,strAttr)

		End If
	Next
End Sub


'::Color Brightness Function
'ex.     ColorBrightness("#FF0000",25,1)	'% Darkness   (of Color Range)
'ex.     ColorBrightness("#FF0000",75,1)	'% Brightness (of Color Range)
'ex.     ColorBrightness("#FF0000",-25,2)	'% Darker     (than Color)
'ex.     ColorBrightness("#FF0000",75,2)	'% Brighter   (than Color)
Function ColorBrightness(strRGB,intBrite,intMode)
	strRGB = Replace(strRGB,"#","")
	
	If Len(strRGB) < 6 Then 'ABD = AABBDD
		str1 = Mid(strRGB,1,1)
		str2 = Mid(strRGB,2,1)
		str3 = Mid(strRGB,3,1)

		strRGB = str1 & str1 & str2 & str2 & str3 & str3
	End If

	r = CInt("&h" & Mid(strRGB,1,2))
	g = CInt("&h" & Mid(strRGB,3,2))
	b = CInt("&h" & Mid(strRGB,5,2))
	arrHSL = RGBtoHSL(r, g, b)
	dblOld = CDbl(arrHSL(2))

	Select Case intMode
		Case 1 '% of Color Range
			dblNew = intBrite/100
		Case 2 '% than Original Color
			dblNew = dblOld + (intBrite/100)
	End Select

	If dblNew > 1 Then
		dblNew = 1
	ElseIf dblNew < 0 Then
		dblNew = 0
	End If

	arrRGB = HSLtoRGB(arrHSL(0), arrHSL(1), dblNew)

	ColorBrightness = "#" & HEXtoDEC(arrRGB(0)) & HEXtoDEC(arrRGB(1)) & HEXtoDEC(arrRGB(2))
End Function


'::RGB to HSL Function
Function RGBtoHSL(r,g,b)
	r = CDbl(r/255)
	g = CDbl(g/255)
	b = CDbl(b/255)

	max = CDbl(MaxCalc(r & "," & g & "," & b))
	min = CDbl(MinCalc(r & "," & g & "," & b))

	h = CDbl((max + min) / 2)
	s = CDbl((max + min) / 2)
	l = CDbl((max + min) / 2)

	If max = min Then
		h = 0
		s = 0
	Else
		d = max - min
		s = IIf(l > 0.5, d / (2 - max - min), d / (max + min))
		Select Case CStr(max)
			Case CStr(r)
				h = (g - b) / d + (IIf(g < b, 6, 0))
			Case CStr(g)
				h = (b - r) / d + 2
            Case CStr(b)
				h = (r - g) / d + 4
		End Select
		h = h / 6
	End If

	RGBtoHSL = Split(h & "," & s & "," & l, ",")
End Function


'::HSL to RGB Function
Function HSLtoRGB(h,s,l)
	If s = 0 Then
		r = l
		g = l
		b = l
	Else
		q = IIf(l < 0.5, l * (1 + s), l + s - l * s)
		p = 2 * l - q
		r = HUEtoRGB(p, q, h + 1/3)
		g = HUEtoRGB(p, q, h)
		b = HUEtoRGB(p, q, h - 1/3)
	End If

	HSLtoRGB = Split(r * 255 & "," & g * 255 & "," & b * 255, ",")
End Function


'::Hue to RGB Function
Function HUEtoRGB(p,q,t)
	If CDbl(t) < 0 Then t = t + 1
	If CDbl(t) > 1 Then t = t - 1

	If CDbl(t) < (1/6) Then
		HUEtoRGB = p + (q - p) * 6 * t
		Exit Function
	End If

	If CDbl(t) < (1/2) Then
		HUEtoRGB = q
		Exit Function
	End If

	If CDbl(t) < (2/3) Then
		HUEtoRGB = p + (q - p) * (2/3 - t) * 6
		Exit Function
	End If

	HUEtoRGB = p
End Function


'::Hex to Decimal Function
Function HEXtoDEC(d)
	h = Hex(Round(d,0))
	h = Right(String(2,"0") & h,2)
    HEXtoDEC = h
End Function


'::Max Function
Function MaxCalc(valList)
	valList = Split(valList,",")
	b = 0
	For v = 0 To UBound(valList)
		a = valList(v)
		If CDbl(a) > CDbl(b) Then b = a
	Next
	MaxCalc = b
End Function


'::Min Function
Function MinCalc(valList)
	valList = Split(valList,",")
	For v = 0 To UBound(valList)
		a = valList(v)
		If b = "" Then b = a
		If CDbl(a) < CDbl(b) AND b <> "" Then b = a
	Next
	MinCalc = b
End Function


'::IIf Emulation Function
Function IIf(condition,conTrue,conFalse)
	If (condition) Then
		IIf = conTrue
	Else
		IIf = conFalse
	End If
End Function


'::Image Size Function (FSO)
'blnImg = ImgInfo(strPath, intImgWidth, intImgHeight, intImgColors, strImgType)
'Response.Write intImgWidth & ", " & intImgHeight & ", " & intImgColors & ", " & strImgType
Function ImgInfo(flnm, width, height, depth, strImageType)
	Dim strPNG
	Dim strGIf
	Dim strBMP
	Dim strType

	'If NOT InStr(flnm,":") Then flnm = Server.MapPath("/" & flnm)

	strType = ""
	strImageType = "(unknown)"
	ImgInfo = False
	strPNG = Chr(137) & Chr(80) & Chr(78)
	strGIf = "GIf"
	strBMP = Chr(66) & Chr(77)
	strType = GetBytes(flnm, 0, 3)

	If strType = strGIf Then 'GIf
		strImageType = "GIf"
		Width = lngConvert(GetBytes(flnm, 7, 2))
		Height = lngConvert(GetBytes(flnm, 9, 2))
		Depth = 2 ^ ((Asc(GetBytes(flnm, 11, 1)) and 7) + 1)
		ImgInfo = True

	ElseIf Left(strType, 2) = strBMP Then 'BMP
		strImageType = "BMP"
		Width = lngConvert(GetBytes(flnm, 19, 2))
		Height = lngConvert(GetBytes(flnm, 23, 2))
		Depth = 2 ^ (Asc(GetBytes(flnm, 29, 1)))
		ImgInfo = True

	ElseIf strType = strPNG Then 'PNG
		strImageType = "PNG"
		Width = lngConvert2(GetBytes(flnm, 19, 2))
		Height = lngConvert2(GetBytes(flnm, 23, 2))
		Depth = getBytes(flnm, 25, 2)
		Select Case Asc(Right(Depth,1))
			Case 0
				Depth = 2 ^ (Asc(Left(Depth, 1)))
				ImgInfo = True
			Case 2
				Depth = 2 ^ (Asc(Left(Depth, 1)) * 3)
				ImgInfo = True
			Case 3
				Depth = 2 ^ (Asc(Left(Depth, 1)))  '8'
				ImgInfo = True
			Case 4
				Depth = 2 ^ (Asc(Left(Depth, 1)) * 2)
				ImgInfo = True
			Case 6
				Depth = 2 ^ (Asc(Left(Depth, 1)) * 4)
				ImgInfo = True
			Case Else
				Depth = -1
		End Select

	Else
		strBuff = GetBytes(flnm, 0, -1) 'Get all bytes from file
		lngSize = len(strBuff)
		flgFound = 0

		strTarget = Chr(255) & Chr(216) & Chr(255)
		flgFound = instr(strBuff, strTarget)

		If flgFound = 0 Then
			Exit Function
		End If

		strImageType = "JPG"
		lngPos = flgFound + 2
		ExitLoop = false

		Do while ExitLoop = False and lngPos < lngSize
			Do while Asc(Mid(strBuff, lngPos, 1)) = 255 and lngPos < lngSize
				lngPos = lngPos + 1
			Loop
			If Asc(Mid(strBuff, lngPos, 1)) < 192 or Asc(Mid(strBuff, lngPos, 1)) > 195 Then
				lngMarkerSize = lngConvert2(Mid(strBuff, lngPos + 1, 2))
				lngPos = lngPos + lngMarkerSize  + 1
			Else
				ExitLoop = True
			End If
		Loop

		If ExitLoop = False Then
			Width = -1
			Height = -1
			Depth = -1
		Else
			Height = lngConvert2(Mid(strBuff, lngPos + 4, 2))
			Width = lngConvert2(Mid(strBuff, lngPos + 6, 2))
			Depth = 2 ^ (Asc(Mid(strBuff, lngPos + 8, 1)) * 8)
			ImgInfo = True
		End If
	End If
End Function

Function GetBytes(flnm, offset, bytes)
	Dim objFSO
	Dim objFTemp
	Dim objTextStream
	Dim lngSize
	On Error Resume Next

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	Set objFTemp = objFSO.GetFile(flnm)
	lngSize = objFTemp.Size
	set objFTemp = Nothing

	fsoForReading = 1
	Set objTextStream = objFSO.OpenTextFile(flnm, fsoForReading)

	If offset > 0 Then
		strBuff = objTextStream.Read(offset - 1)
	End If

	If bytes = -1 Then
		GetBytes = objTextStream.Read(lngSize)
	Else
		GetBytes = objTextStream.Read(bytes)
	End If

	objTextStream.Close
	set objTextStream = Nothing
	set objFSO = Nothing
End Function

Function lngConvert(strTemp)
	lngConvert = CLng(Asc(Left(strTemp, 1)) + ((Asc(Right(strTemp, 1)) * 256)))
End Function

Function lngConvert2(strTemp)
	lngConvert2 = CLng(Asc(Right(strTemp, 1)) + ((Asc(Left(strTemp, 1)) * 256)))
End Function


'::PCase Function (capitalize first letters)
Function PCase(strInput)
	Dim iPosition 
	Dim iSpace    
	Dim strOutput 
	iPosition = 1
	Do While InStr(iPosition, strInput, " ", 1) <> 0
		iSpace = InStr(iPosition, strInput, " ", 1)
		strOutput = strOutput & UCase(Mid(strInput, iPosition, 1))
		strOutput = strOutput & LCase(Mid(strInput, iPosition + 1, iSpace - iPosition))
		iPosition = iSpace + 1
	Loop
	strOutput = strOutput & UCase(Mid(strInput, iPosition, 1))
	strOutput = strOutput & LCase(Mid(strInput, iPosition + 1))
	PCase = strOutput
End Function