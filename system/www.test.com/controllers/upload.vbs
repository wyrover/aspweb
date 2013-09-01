Import "system/libraries/clsupfile.vbs"	
Import "system/libraries/clsfileobj.vbs"	
'Import "system/libraries/clsThumb.vbs
Import "system/libraries/upfile_class.vbs"	


Class Upload
	Function index()
		Select case segment(3)
			case "2"
				d("form") = Picture_UpPhoto()
			case else
				d("form") = Picture_UpPhoto()
		End Select
		t.Load "manage/upload.htm", d
	End Function

	Function Picture_UpPhoto()
		Dim Path, InstallDir, DateDir, sb		
		'Path = KSCMS.ReturnChannelUpFilesDir(2)
		Path = "\attachments\" & Application_PATH
		DateDir = Year(Now()) & Right("0" & Month(Now()), 2) & "/"
		Path = Path & "/" & DateDir
		'Call KSCMS.CreateListFolder(Path)
		Set sb = New StringBuilder
		sb.Append "<div align=""center"">"
		sb.Append "  <table width=""95%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
		sb.Append "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""?/upload/UpFileSave"">"
		sb.Append "      <tr>"
		sb.Append "        <td width=""82%"" valign=""top"">"
		sb.Append "          <div align=""center""> <br>"
		sb.Append "            <table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		sb.Append "              <tr>"
		sb.Append "                <td width=""50%"" height=""50""> &nbsp;&nbsp;设定上传图片数量"
		sb.Append "                  <input name=""UpFileNum"" class=""upfile"" type=""text"" value=""10"" size=""10"">"
		sb.Append "                <input type=""button"" name=""Submit42"" class='button' value=""确定设定"" onClick=""ChooseOption();""></td>"
		sb.Append "                <td width=""50%"" id='ss'>&nbsp;</td>"
		sb.Append "              </tr>"
		sb.Append "              <tr>"
		sb.Append "                <td height=""30"" colspan=""2"" id=""FilesList""> </td>"
		sb.Append "              </tr>"
		sb.Append "            </table>"
		sb.Append "        </div></td>"
		sb.Append "        <td width=""18%"" valign=""top""><br>"
		sb.Append "          <br>"
		sb.Append "          <br><br>"
		sb.Append "          <input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
		sb.Append "添加水印<br>"
		sb.Append "<br>"
		sb.Append " <fieldset style=""width:100%;"">"
		sb.Append "          <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		sb.Append "            <tr>"
		sb.Append "              <td height=""20"">"
		sb.Append "                <div align=""center"">命名规则</div></td>"
		sb.Append "            </tr>"
		sb.Append "            <tr>"
		sb.Append "              <td height=""20"">"
		sb.Append "                <div align=""left"">"
		sb.Append "                  <input type=""radio"" name=""AutoReName"" value=""0"">"
		sb.Append "                  原名称不变</div></td>"
		sb.Append "            </tr>"
		sb.Append "            <tr>"
		sb.Append "              <td height=""20"">"
		sb.Append "                <div align=""left"">"
		sb.Append "                  <input type=""radio"" name=""AutoReName"" value=""1"">"
		sb.Append "                  &quot; 副件&quot;+文件名</div></td>"
		sb.Append "            </tr>"
		sb.Append "            <tr>"
		sb.Append "              <td height=""20"">"
		sb.Append "                <div align=""left"">"
		sb.Append "                  <input type=""radio"" name=""AutoReName"" value=""2"">"
		sb.Append "                  随机数+扩展名</div></td>"
		sb.Append "            </tr>"
		sb.Append "            <tr>"
		sb.Append "              <td height=""20""><input type=""radio"" name=""AutoReName"" value=""3"">"
		sb.Append "              随机数+文件名</td>"
		sb.Append "            </tr>"
		sb.Append "            <tr>"
		sb.Append "              <td height=""20"">"
		sb.Append "                <div align=""left"">"
		sb.Append "                  <input name=""AutoReName"" type=""radio"" value=""4"" checked>"
		sb.Append "                  20060101121022</div></td>"
		sb.Append "            </tr>"
		sb.Append "          </table>"
		sb.Append "        </fieldset></td>"
		sb.Append "      </tr>"
		sb.Append "      <tr>"
		sb.Append "        <td  colspan=""2"" align='center'>"
		sb.Append "                  <input type=""submit"" id=""BtnSubmit""  class='button' name=""Submit"" value=""开始上传"">"
		sb.Append "                  <input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
		sb.Append "                  <input name=""UpLoadFrom"" value=""21"" type=""hidden"" id=""UpLoadFrom"">"
		sb.Append "                  <input type=""reset"" id=""ResetForm"" class='button' name=""Submit3"" value="" 重 填 "">"
		sb.Append "        </td>"
		sb.Append "      </tr>"
		sb.Append "    </form>"
		sb.Append "  </table>"
		sb.Append "</div>"
		sb.Append "<script language=""JavaScript""> " & vbCrLf
		sb.Append "function ChooseOption()" & vbCrLf
		sb.Append "{"
		sb.Append "  var UpFileNum = document.all.UpFileNum.value;" & vbCrLf
		sb.Append "  if (UpFileNum=='') " & vbCrLf
		sb.Append "    UpFileNum=10;" & vbCrLf
		sb.Append "  var k,i,Optionstr,SelectOptionstr,n=0;" & vbCrLf
		sb.Append "      Optionstr = '<table width=""100%"" border=""0"" cellspacing=""5"" cellpadding=""0"">';" & vbCrLf
		sb.Append "  for(k=0;k<(UpFileNum/2);k++)" & vbCrLf
		sb.Append "   { Optionstr = Optionstr+'<tr>';" & vbCrLf
		sb.Append "    for (i=0;i<2;i++)" & vbCrLf
		sb.Append "      { n=n+1;" & vbCrLf
		sb.Append "       Optionstr = Optionstr+'<td>&nbsp;图&nbsp;片&nbsp;'+n+'</td><td>&nbsp;<input type=""file"" accept=""html"" size=""20"" class=""upfile"" name=""File'+n+'"">&nbsp;</td>';" & vbCrLf
		sb.Append "        if (n==UpFileNum) break;" & vbCrLf
		sb.Append "       }" & vbCrLf
		sb.Append "      while (i <= 2)" & vbCrLf
		sb.Append "      {" & vbCrLf
		sb.Append "      Optionstr = Optionstr+'<td width=""50%"">&nbsp; </td>';" & vbCrLf
		sb.Append "      i++;" & vbCrLf
		sb.Append "      }" & vbCrLf
		sb.Append "      Optionstr = Optionstr+'</tr>'" & vbCrLf
		sb.Append "  }" & vbCrLf
		sb.Append "    Optionstr = Optionstr+'</table>';" & vbCrLf
		sb.Append "    document.all.FilesList.innerHTML = Optionstr;" & vbCrLf
		sb.Append "    SelectOptionstr='设定<select class=""upfile"" name=""DefaultUrl"">'" & vbCrLf
		sb.Append " for(i=1;i<=UpFileNum;++i)" & vbCrLf
		sb.Append "  {" & vbCrLf
		sb.Append "   SelectOptionstr=SelectOptionstr+'<option value=""'+eval(i)+'"">第'+eval(i)+'张图片</option>'" & vbCrLf
		sb.Append "  }" & vbCrLf
		sb.Append "   SelectOptionstr=SelectOptionstr+'</select>为缩略图(系统自动生成)'" & vbCrLf
		sb.Append "   document.all.ss.innerHTML=SelectOptionstr;" & vbCrLf
		sb.Append " }" & vbCrLf
		sb.Append "ChooseOption();" & vbCrLf
		sb.Append "</script>" & vbCrLf
		Picture_UpPhoto = sb.ToString()
	End Function

	Function UpFileSave()
		Dim FilePath, MaxFileSize, AllowFileExtStr, AutoReName, RsConfigObj
		Dim FormName, Path, UpLoadFrom, TempFileStr, FormPath, ThumbFileName, ThumbPathFileName
		Dim UpFileObj, FsoObjName, AddWaterFlag, T, CurrNum, CreateThumbsFlag
		Dim DefaultThumb    '设定第几张为缩略图
		Dim ReturnValue

		Response.Write("<style type='text/css'>" & vbcrlf)
		Response.Write("<!--" & vbcrlf)
		Response.Write("body {background:#EEF8FE;" & vbcrlf)
		Response.Write("	margin-left: 0px;" & vbcrlf)
		Response.Write("	margin-top: 0px;" & vbcrlf)
		Response.Write("}" & vbcrlf)
		Response.Write("-->" & vbcrlf)
		Response.Write("</style>" & vbcrlf)

		Set UpFileObj = New UpFileClass
		UpFileObj.GetData
		FormPath = UpFileObj.Form("Path") 
		
		Call CreateFolder(FormPath) 
		FilePath = Server.MapPath(FormPath) & "\"

		AutoReName = UpFileObj.Form("AutoRename")
		UpLoadFrom = UpFileObj.Form("UpLoadFrom")        
		'0--通用对话框 2-- 图片中心上传 31--下载中心缩略图 32--下载中心文件 41--动漫中心缩略图 42--动漫中心的动漫文件
		IF UpLoadFrom = "" then
		  UpLoadFrom = 0
		End IF

		CurrNum = 0
		CreateThumbsFlag = False
		DefaultThumb = UpFileObj.Form("DefaultUrl")
		If DefaultThumb = "" then DefaultThumb = 0
		AddWaterFlag = UpFileObj.Form("AddWaterFlag")
		If AddWaterFlag <> "1" Then	'生成是否要添加水印标记
			AddWaterFlag = "0"
		End if

		MaxFileSize = 1024
		AllowFileExtStr = "jpg|gif|png"
		

		ReturnValue = CheckUpFile(FilePath,MaxFileSize,AllowFileExtStr,AutoReName, UpFileObj)

		If ReturnValue <> "" Then
			Response.Write("<script language=""JavaScript"">")
			Response.Write("alert('" & ReturnValue & "');")
			Response.Write("history.back(-1);")
			Response.Write("</script>")
		Else  
			Response.Write("<script language=""JavaScript"">")
			Response.Write("parent.SetPicUrlByUpLoad('" & TempFileStr &  "','" & ThumbPathFileName & "');")
			Response.Write("document.write('<br><br><div align=center><font size=2>图片上传成功！</font></div>');")
			Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'1; url=../Admin_UpFileForm.asp?ChannelID=2\'>');")
			Response.Write("</script>")
		End If
	End Function

	Function CheckUpFile(Path,FileSize,AllowExtStr,AutoReName, UpFileObj)
		Dim ErrStr, NoUpFileTF, FsoObj, FileName, FileExtName, FileContent, SameFileExistTF
		NoUpFileTF = True
		ErrStr = ""
		Set FsoObj = Server.CreateObject("Scripting.FileSystemObject")
		
		For Each FormName in UpFileObj.File
			SameFileExistTF = False			

			FileName = UpFileObj.File(FormName).FileName
			FileExtName = UpFileObj.File(FormName).FileExt
			FileContent = UpFileObj.File(FormName).FileData
			
			'是否存在重名文件
			If UpFileObj.File(FormName).FileSize > 1 Then
				
				NoUpFileTF = False
				ErrStr = ""
				If UpFileObj.File(FormName).FileSize > CLng(FileSize)*1024 Then
					ErrStr = ErrStr & FileName & "文件上传失败\n超过了限制，最大只能上传" & FileSize & "K的文件\n"
				End If

				If AutoRename = "0" Then
					If FsoObj.FileExists(Path & FileName) = True  Then
						ErrStr = ErrStr & FileName & "文件上传失败,存在同名文件\n"
					Else
						SameFileExistTF = True
					End If
				Else
					SameFileExistTF = True
				End If



				If CheckFileType(AllowExtStr,FileExtName) = False Then
					ErrStr = ErrStr & FileName & "文件上传失败,文件类型不允许\n允许的类型有" + AllowExtStr + "\n"
				End If

				

				If ErrStr = "" Then
					If SameFileExistTF = True Then
						SaveFile Path, FormName, AutoReName, UpFileObj
					Else
						SaveFile Path, FormName, "", UpFileObj
					End If
				Else
					CheckUpFile = CheckUpFile & ErrStr
				End If

			End If

		Next

		Set FsoObj = Nothing

		If NoUpFileTF = True Then
			CheckUpFile = "没有上传文件"
		End If
	End Function

	Function CheckFileType(AllowExtStr, FileExtName)
		Dim i,AllowArray
		AllowArray = Split(AllowExtStr,"|")
		FileExtName = LCase(FileExtName)
		CheckFileType = False
		For i = LBound(AllowArray) to UBound(AllowArray)
			if LCase(AllowArray(i)) = LCase(FileExtName) then
				CheckFileType = True
			end if
		Next
		if FileExtName="asp" or FileExtName="asa" or FileExtName="aspx" then
			CheckFileType = False
		end if
	End Function

	Function SaveFile(FilePath,FormNameItem,AutoNameType, UpFileObj)

		Dim FileName, FileExtName, FileContent, FormName, RandomFigure, n, RndStr, T
		'Set T = New Thumb
		Randomize 
		n = 2* Rnd + 10
		RndStr = MakeRandom(n)
		
		RandomFigure = CStr(Int((99999 * Rnd) + 1))
		FileName = UpFileObj.File(FormNameItem).FileName
		FileExtName = UpFileObj.File(FormNameItem).FileExt		
		FileContent = UpFileObj.File(FormNameItem).FileData

		
		Select case AutoNameType 
		  Case "1"
			FileName= "副件" & FileName
		  Case "2"
			FileName= RndStr&"."&FileExtName
		  Case "3"
			FileName= RndStr & FileName
		  Case "4"
			FileName= Year(Now())&Right("0"&Month(Now()),2)&Right("0"&Day(Now()),2)&Right("0"&Hour(Now()),2)&Right("0"&Minute(Now()),2)&Right("0"&Second(Now()),2)&RandomFigure&"."&FileExtName
		  Case Else
			FileName=FileName
		End Select

		UpFileObj.File(FormNameItem).SaveToFile FilePath  & FileName
		TempFileStr = TempFileStr & FormPath & FileName & "|"

		If AddWaterFlag = "1" Then   '在保存好的图片上添加水印			
			call T.AddWaterMark(FilePath  & FileName)
		End if

		CurrNum=CurrNum+1

		If CreateThumbsFlag = True and  CInt(CurrNum) = CInt(DefaultThumb) Then
			ThumbFileName=split(FileName,".")(0)&"_S."&FileExtName
			'call T.CreateThumbs(FilePath & FileName,FilePath & ThumbFileName)
			 '取得缩略图地址
			ThumbPathFileName = FormPath & ThumbFileName
		End if

	

	
	End Function


	Function index2()
		t.Load "manage/upload2.htm", d
	End Function

	Function index3()
		If segment(3) = "img" Then
			d("form") = "<iframe id=d_file frameborder=0 src=""?/upload/upfile/img"" width=""100%"" height=""45"" scrolling=no></iframe>"
		Else 
			d("form") = "<iframe id=d_file frameborder=0 src=""?/upload/upfile/soft"" width=""100%"" height=""45"" scrolling=no></iframe>"
		End If
		t.Load "manage/upfileForm.htm", d
	End Function

	Function upfile()
		Dim upType,sAllowExt,sUploadDir,nAllowSize,message,sPathFileName

		If segment(3) <> "0" Then
			upType = Lcase(segment(3))
			Select case upType
				case "img"
					nAllowSize = config("upload_filesize") 
					sAllowExt = config("upload_file_ext")
					sUploadDir = config("upload_img_dir") 
				case "soft"
					nAllowSize = config("upload_filesize") 
					sAllowExt = config("upload_file_ext")
					sUploadDir = config("upload_img_dir") 
				case else
					'Call ErrBox("非法的传值参数!")
			End Select
		End If

		If segment(4) = "post" Then
			Dim oUpload, oFile,oFloder
			' 建立上传对象
			Set oUpload = New upfile_class
			' 取得上传数据,限制最大上传
			oUpload.GetData(config("upload_filesize") *1024)

			If oUpload.Err > 0 Then
			Select Case oUpload.Err
				Case 1
					Call MsgBox2("请选择有效的上传文件", 0, "0")				
				Case 2
					Call MsgBox2("你上传的文件总大小超出了最大限制:" & config("upload_filesize")  & "KB", 0, "0")					
				End Select
				Response.End
			End If

			Set oFile = oUpload.File("uploadfile")
			sFileExt = LCase(oFile.FileExt)
			Call CheckValidExt(sFileExt)
			sOriginalFileName = oFile.FileName
			sSaveFileName = GetRndFileName(sFileExt)
			oFloder = year(now) & month(now) & day(now)
			CreateFolder(config("archives") & "\images\" & oFloder)
			sUploadDir = sUploadDir & oFloder
			oFile.SaveToFile sUploadDir & "\" & sSaveFileName
			sPathFileName = Replace(config("archives") & "images\" & oFloder & "\" & sSaveFileName, "\", "/")
			Set oFile = Nothing
			Set oUpload = Nothing
			Response.Write "<script>parent.UploadSaved('" & sPathFileName & "');</script>"
		End If
		Response.Write "<html>" & vbCrlf
		Response.Write "<head>" & vbCrlf
		Response.Write "<script language=""javascript"">" & vbCrlf
		Response.Write "  function returnValue(value)" & vbCrlf
		Response.Write "  {" & vbCrlf
		Response.Write "  parent.window.returnValue = value;" & vbCrlf
		Response.Write "  parent.window.close();" & vbCrlf
		Response.Write "  }" & vbCrlf
		Response.Write "</script>" & vbCrlf
		Response.Write "</head>" & vbCrlf
		Response.Write "<body bgcolor=menu>" & vbCrlf
		Response.Write vbCrlf
		Response.Write "<table width=""100%""  border=""0"" cellspacing=""1"" cellpadding=""0"">" & vbCrlf
		Response.Write "  <tr>" & vbCrlf
		Response.Write "    <td>" & vbCrlf
		Response.Write "	<form action=""?/upload/upfile/img/post"" method=""post"" enctype=""multipart/form-data"" name=""form1"" class=""noMarginform"" onSubmit=""document.form1.btnSubmit.disabled=true;"">" & vbCrlf
		Response.Write "      <input name=uploadfile type=""file"" size=""25"">" & vbCrlf
		Response.Write "	  <input type=""submit"" name=""btnSubmit"" value=""上传"" onClick=""window.divProcessing.style.display='';"">" & vbCrlf
		Response.Write "    </form></td>" & vbCrlf
		Response.Write "  </tr>" & vbCrlf
		Response.Write "</table>" & vbCrlf
		Response.Write "<div id=divProcessing style=""width:350px;height:30px;position:absolute;left:10px;top:10px;display:none;"">" & vbCrlf
		Response.Write "<table border=0 cellpadding=0 cellspacing=1 bgcolor=""#000000"" width=""100%"" height=""100%""><tr><td bgcolor=""#6B91CF""><marquee align=""middle"" behavior=""alternate"" scrollamount=""5""><font color=#ffffff size=""2"">...文件上传中...请等待...</font></marquee></td></tr></table>" & vbCrlf
		Response.Write "</div>" & vbCrlf
		Response.Write "</body>" & vbCrlf
		Response.Write "</html>" & vbCrlf

	End Function

	Sub CheckValidExt(sExt)
		Dim b, i, aExt
		b = False
		aExt = Split(config("upload_file_ext"), "|")
		For i = 0 To UBound(aExt)
			If LCase(aExt(i)) = sExt Then
				b = True
				Exit For
			End If
		Next
		If b = False Then
			Call MsgBox2("非法的文件格式", 0, "0")					
		End If
	End Sub

	Function GetRndFileName(sExt)
		Dim sRnd
		Randomize
		sRnd = Int(900 * Rnd) + 100
		GetRndFileName = year(now) & month(now) & day(now) & hour(now) & minute(now) & second(now) & sRnd & "." & sExt
	End Function


End Class

