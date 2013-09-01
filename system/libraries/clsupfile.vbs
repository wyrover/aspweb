'-----------------------------------------------------------------------
'--- 上传处理类模块
'--- Copyright (c) 2004 Aspsky, Inc.
'--- Mail: Sunwin@artbbs.net   http://www.aspsky.net
'--- 2004-12-18
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
'-- InceptFileType	: 设置上传类型属性 (以逗号分隔多个文件类型) String
'-- MaxSize			: 设置上传文件大小上限 (单位：kb) Long
'-- InceptMaxFile	: 设置一次上传文件最大个数 Long
'-- UploadPath		: 设置保存的目录相对路径 String
'-- UploadType		: 设置上传组件类型 （0=无组件上传类，1=Aspupload3.0 ,2=SA-FileUp 4.0 ,3=DvFile.Upload V1.0）
'-- SaveUpFile		: 执行上传
'-- GetBinary		: 设置上传是否返回文件数据流  Bloon值 : True/False
'-- ChkSessionName	: 设置SESSION名，防止重复提交，SESSION名与提交的表单名要一致。
'-- IsReName        : 是否重命名，0为原文件名，1重新按日期命名。（梅傲风）
'-- RName设置文件名	: 定义文件名前缀 (如默认生成的文件名为200412230402587123.jpg
'									设置：RName="PRE_",生成的文件名为：PRE_200412230402587123.jpg)
'-----------------------------------------------------------------------
'-- 设置图片组件属性
'-- PreviewType		: 设置组件(0=CreatePreviewImage组件，1=AspJpegV1.2 ,2=SoftArtisans ImgWriter V1.21)
'-- PreviewImageWidth	: 设置预览图片宽度
'-- PreviewImageHeight	: 设置预览图片高度
'-- DrawImageWidth	: 设置水印图片或文字区域宽度
'-- DrawImageHeight	: 设置水印图片或文字区域高度
'-- DrawGraph		: 设置水印图片或文字区域透明度
'-- DrawFontColor	: 设置水印文字颜色
'-- DrawFontFamily	: 设置水印文字字体格式
'-- DrawFontSize	: 设置水印文字字体大小
'-- DrawFontBold	: 设置水印文字是否粗体
'-- DrawInfo		: 设置水印文字信息或图片信息
'-- DrawType		: 设置加载水印模式：0=不加载水印 ，1=加载水印文字 ，2=加载水印图片
'-- DrawXYType		: 图片添加水印LOGO位置坐标："0" =左上，"1"=左下,"2"=居中,"3"=右上,"4"=右下
'-- DrawSizeType	: 生成预览图片大小规则："0"=固定缩小，"1"=等比例缩小
'-----------------------------------------------------------------------
'-- 获取上传信息
'-- ObjName			: 采用的组件名称
'-- Count			: 上传文件总数
'-- CountSize		: 上传总大小字节数
'-- ErrCodes		: 错误NUMBER (默认为0)
'-- Description		: 错误描述
'-----------------------------------------------------------------------
'-- CreateView Imagename,TempFilename,FileExt
'	创建预览图片过程: 原始文件的相对路径,生成预览文件相对路径,原文件后缀
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
'-- 获取文件对象属性 : UploadFiles
'-- FormName		: 表单名称
'-- OldFileName		: 原文件名称    （By梅傲风）
'-- FileName		: 生成的文件名称
'-- FilePath		: 保存文件的相对路径
'-- FileSize		: 文件大小
'-- FileContentType	: ContentType文件类型
'-- FileType		: 0=其它,1=图片,2=FLASH,3=音乐,4=电影
'-- FileData		: 文件数据流 (若组件不支持直接获取，则返回Null)
'-- FileExt			: 文件后缀
'-- FileWidth		: 图片/Flash文件宽度	（其他文件默认=-1）
'-- FileHeight		: 图片/Flash文件高度	（其他文件默认=-1）
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
'-- 获取表单对象属性 : UploadForms
'-- Count			: 表单数
'-- key				: 表单内容
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------

Dim oUpFileStream

Class UpFile_Cls
	Private UploadObj,ImageObj
	Private FilePath,InceptFile,FileMaxSize,MaxFile,Upload_Type,FileInfo,IsBinary,SessionName
	Private Preview_Type,View_ImageWidth,View_ImageHeight,Draw_ImageWidth,Draw_ImageHeight,Draw_Graph
	Private Draw_FontColor,Draw_FontFamily,Draw_FontSize,Draw_FontBold,Draw_Info,Draw_Type,Draw_XYType,Draw_SizeType
	Private IsReName_Str,RName_Str,Transition_Color
	Public ErrCodes,ObjName,UploadFiles,UploadForms,Count,CountSize
	Public ProgressID_Str
	'-----------------------------------------------------------------------------------
	'初始化类
	'-----------------------------------------------------------------------------------
	Private Sub Class_Initialize
		SessionName = Empty
		IsBinary = False
		ErrCodes = 0
		Count = 0
		CountSize = 0
		FilePath = "./"
		InceptFile = ""
		FileMaxSize = -1
		MaxFile = 1
		Upload_Type = -1
		Preview_Type = 999
		IsReName_Str=1   'By 梅傲风
		ObjName = "未知组件"
		View_ImageWidth = 0
		View_ImageHeight = 0
		Draw_FontColor	= &H000000
		Draw_FontFamily	= "Arial"
		Draw_FontSize	= 10
		Draw_FontBold	= False
		Draw_Info		= "Www.Aspoo.CN"
		Draw_Type		= -1
		Set UploadFiles = Server.CreateObject ("Scripting.Dictionary")
		Set UploadForms = Server.CreateObject ("Scripting.Dictionary")
		UploadFiles.CompareMode = 1
		UploadForms.CompareMode = 1
	End Sub

	'-----------------------------------------------------------------------------------
	'销毁类
	'-----------------------------------------------------------------------------------
	Private Sub Class_Terminate
		If IsObject(UploadObj) Then
			Set UploadObj = Nothing
		End If
		If IsObject(ImageObj) Then
			Set ImageObj = Nothing
		End If
		UploadFiles.RemoveAll
		UploadForms.RemoveAll
		Set UploadForms = Nothing
		Set UploadFiles = Nothing
	End Sub

	'-----------------------------------------------------------------------------------
	'设置上传是否返回文件数据流
	'-----------------------------------------------------------------------------------
	Public Property Let GetBinary(Byval Values)
		IsBinary = Values
	End Property

	'-----------------------------------------------------------------------------------
	'设置上传类型属性 (以逗号分隔多个文件类型)
	'-----------------------------------------------------------------------------------
	Public Property Let InceptFileType(Byval Values)
		InceptFile = Lcase(Values)
	End Property

	'-----------------------------------------------------------------------------------
	'设置上传类型属性 (以逗号分隔多个文件类型)
	'-----------------------------------------------------------------------------------
	Public Property Let ChkSessionName(Byval Values)
		SessionName = Values
	End Property

	'-----------------------------------------------------------------------------------
	'设置上传文件大小上限 (单位：kb)
	'-----------------------------------------------------------------------------------
	Public Property Let MaxSize(Byval Values)
		FileMaxSize = ChkNumeric(Values) * 1024
	End Property
	Public Property Get MaxSize
		MaxSize = FileMaxSize
	End Property

	'-----------------------------------------------------------------------------------
	'设置每次上传文件上限
	'-----------------------------------------------------------------------------------
	Public Property Let InceptMaxFile(Byval Values)
		MaxFile = ChkNumeric(Values)
	End Property

	'-----------------------------------------------------------------------------------
	'设置上传目录路径
	'-----------------------------------------------------------------------------------
	Public Property Let UploadPath(Byval Path)
		FilePath = Replace(Path,Chr(0),"")
		If Right(FilePath,1)<>"/" Then FilePath = FilePath & "/"
	End Property

	Public Property Get UploadPath
		UploadPath = FilePath
	End Property

	'-----------------------------------------------------------------------------------
	'获取错误信息
	'-----------------------------------------------------------------------------------
	Public Property Get Description
		Select Case ErrCodes
			Case 1 : Description = "不支持 " & ObjName & "，服务器可能未安装该组件。"
			Case 2 : Description = "暂未选择上传组件！"
			Case 3 : Description = "请先选择你要上传的文件!"
			Case 4 : Description = "文件大小超过了限制 " & (FileMaxSize\1024) & "KB!"
			Case 5 : Description = "文件类型不正确，只允许上传类型为 "&InceptFile&" 的文件!"
			Case 6 : Description = "已达到上传数的上限！"
			Case 7 : Description = "请不要重复提交！"
			Case Else
				Description = Empty
		End Select
	End Property

	'-----------------------------------------------------------------------------------
	'是否重命名（梅傲风）
	'-----------------------------------------------------------------------------------
	Public Property Let IsReName(Byval Values)
		IsReName_Str = ChkNumeric(Values)
	End Property

	'-----------------------------------------------------------------------------------
	'设置文件名前缀
	'-----------------------------------------------------------------------------------
	Public Property Let RName(Byval Values)
		RName_Str = Values
	End Property

	'-----------------------------------------------------------------------------------
	'设置上传组件属性
	'-----------------------------------------------------------------------------------
	Public Property Let UploadType(Byval Types)
		Upload_Type = Types
		If Upload_Type = "" or Not IsNumeric(Upload_Type) Then
			Upload_Type = -1
		End If
	End Property

	Public Property Let ProgressID(Byval Values)
		ProgressID_Str = Values
	End Property
	

	'-----------------------------------------------------------------------------------
	'设置上传图片组件属性
	'-----------------------------------------------------------------------------------
	Public Property Let PreviewType(Byval Types)
		Preview_Type = Types
		'On Error Resume Next
		If Preview_Type = "" or Not IsNumeric(Preview_Type) Then
			Preview_Type = 999
		'Else
		'	If PreviewType <> 999 Then
		'		Select Case Preview_Type
		'			Case 0
					'---------------------CreatePreviewImage---------------
		'				ObjName = "CreatePreviewImage组件"
		'				Set ImageObj = Server.CreateObject("CreatePreviewImage.cGvbox")
		'			Case 1
					'---------------------AspJpegV1.2---------------
		'				ObjName = "AspJpegV1.2组件"
		'				Set ImageObj = Server.CreateObject("Persits.Jpeg")
		'			Case 2
					'---------------------SoftArtisans ImgWriter V1.21---------------
		'				ObjName = "SoftArtisans ImgWriter V1.21组件"
		'				Set ImageObj = Server.CreateObject("SoftArtisans.ImageGen")
		'			Case Else
		'				Preview_Type = 999
		'		End Select
		'		If Err.Number<>0 Then
		'			ErrCodes = 1
		'		End If
		'	End If
		End If
	End Property

	Public Property Get PreviewType
		PreviewType = Preview_Type
	End Property

	'-----------------------------------------------------------------------------------
	'设置预览图片宽度属性
	'-----------------------------------------------------------------------------------
	Public Property Let PreviewImageWidth(Byval Values)
		View_ImageWidth = ChkNumeric(Values)
	End Property

	'-----------------------------------------------------------------------------------
	'设置预览图片高度属性
	'-----------------------------------------------------------------------------------
	Public Property Let PreviewImageHeight(Byval Values)
		View_ImageHeight = ChkNumeric(Values)
	End Property

	'-----------------------------------------------------------------------------------
	'设置水印图片或文字区域宽度属性
	'-----------------------------------------------------------------------------------
	Public Property Let DrawImageWidth(Byval Values)
		Draw_ImageWidth = ChkNumeric(Values)
	End Property

	'-----------------------------------------------------------------------------------
	'设置水印图片或文字区域高度属性
	'-----------------------------------------------------------------------------------
	Public Property Let DrawImageHeight(Byval Values)
		Draw_ImageHeight = ChkNumeric(Values)
	End Property

	'-----------------------------------------------------------------------------------
	'设置水印图片或文字区域透明度属性
	'-----------------------------------------------------------------------------------
	Public Property Let DrawGraph(Byval Values)
		If IsNumeric(Values) Then
			Draw_Graph = Formatnumber(Values,2)
		Else
			Draw_Graph = 1
		End If
	End Property

	'-----------------------------------------------------------------------------------
	'设置水印图片透明度去除底色值
	'-----------------------------------------------------------------------------------
	Public Property Let TransitionColor(Byval Values)
		If Values<>"" or Values<>"0" Then
			Transition_Color = Replace(Values,"#","&h")
		End If
	End Property

	'-----------------------------------------------------------------------------------
	'设置水印文字颜色
	'-----------------------------------------------------------------------------------
	Public Property Let DrawFontColor(Byval Values)
		If Values<>"" or Values<>"0" Then
			Draw_FontColor = Replace(Values,"#","&h")
		End If
	End Property

	'-----------------------------------------------------------------------------------
	'设置水印文字字体格式
	'-----------------------------------------------------------------------------------
	Public Property Let DrawFontFamily(Byval Values)
		Draw_FontFamily = Values
	End Property

	'-----------------------------------------------------------------------------------
	'设置水印文字字体大小
	'-----------------------------------------------------------------------------------
	Public Property Let DrawFontSize(Byval Values)
		Draw_FontSize = Values
	End Property

	'-----------------------------------------------------------------------------------
	'设置水印文字是否粗体 Boolean
	'-----------------------------------------------------------------------------------
	Public Property Let DrawFontBold(Byval Values)
		Draw_FontBold = ChkBoolean(Values)
	End Property
	'-----------------------------------------------------------------------------------
	'设置水印文字信息或图片信息
	'-----------------------------------------------------------------------------------
	Public Property Let DrawInfo(Byval Values)
		Draw_Info = Values
	End Property

	'-----------------------------------------------------------------------------------
	'加载模式：0=不加载水印 ，1=加载水印文字 ，2=加载水印图片
	'-----------------------------------------------------------------------------------
	Public Property Let DrawType(Byval Values)
		Draw_Type = ChkNumeric(Values)
	End Property

	'-----------------------------------------------------------------------------------
	'图片添加水印LOGO位置坐标："0" =左上，"1"=左下,"2"=居中,"3"=右上,"4"=右下
	'-----------------------------------------------------------------------------------
	Public Property Let DrawXYType(Byval Values)
		 Draw_XYType = Values
	End Property

	'-----------------------------------------------------------------------------------
	'生成预览图片大小规则："0"=固定缩小，"1"=等比例缩小
	'-----------------------------------------------------------------------------------
	Public Property Let DrawSizeType(Byval Values)
		Draw_SizeType = Values
	End Property

	Private Function ChkNumeric(Byval Values)
		If Values<>"" and Isnumeric(Values) Then
			ChkNumeric = Int(Values)
		Else
			ChkNumeric = 0
		End If
	End Function

	Private Function ChkBoolean(Byval Values)
		If Typename(Values)="Boolean" or IsNumeric(Values) or Lcase(Values)="false" or Lcase(Values)="true" Then
			ChkBoolean = CBool(Values)
		Else
			ChkBoolean = False
		End If
	End Function

	Private Function ChkFileIsExist(strFileName)
		ChkFileIsExist=False
		If Cl.ChkObjInstalled(Trim(Cl.Web_Info(13))) Then
			dim fso
			set fso=CreateObject(Trim(Cl.Web_Info(13)))
			if fso.FileExists(Server.mappath(FilePath & strFileName)) then
				ChkFileIsExist=True
			end if
			set fso=Nothing
		end if
	End Function
	'-----------------------------------------------------------------------------------
	'取得文件名
	'-----------------------------------------------------------------------------------
	Private Function FormatName(Byval OldFileName, Byval FileExt)
		Dim TempStr
		if IsReName_Str=0 then
			TempStr = Left(OldFileName,InStrRev(OldFileName, ".")-1) & "." & FileExt
			If ChkFileIsExist(TempStr) then TempStr = NewFileName & "." & FileExt
		else
			TempStr = NewFileName & "." & FileExt
		end if
		If RName_Str<>"" Then
			TempStr = RName_Str & TempStr
		End If
		FormatName = TempStr
	End Function

	'-----------------------------------------------------------------------------------
	'日期时间定义文件名
	'-----------------------------------------------------------------------------------
	Private Function NewFileName()
		Dim RanNum
		Randomize
		RanNum = Int(90000*rnd)+10000
		NewFileName = Year(now) & Right("0"&Month(now),2) & Right("0"&Day(now),2) & Right("0"&Hour(now),2) & Right("0"&Minute(now),2) & Right("0"&Second(now),2) & RanNum
	End Function
	'-----------------------------------------------------------------------------------
	'格式后缀
	'-----------------------------------------------------------------------------------
	Private Function FixName(Byval UpFileExt)
		If IsEmpty(UpFileExt) Then Exit Function
		FixName = Lcase(UpFileExt)
		FixName = Replace(FixName,Chr(0),"")
		FixName = Replace(FixName,".","")
		FixName = Replace(FixName,"'","")
		FixName = Replace(FixName,"asp","")
		FixName = Replace(FixName,"asa","")
		FixName = Replace(FixName,"aspx","")
		FixName = Replace(FixName,"cer","")
		FixName = Replace(FixName,"cdx","")
		FixName = Replace(FixName,"htr","")
	End Function

	'-----------------------------------------------------------------------------------
	'判断文件类型是否合格
	'-----------------------------------------------------------------------------------
	Private Function CheckFileExt(FileExt)
		Dim Forumupload,i
		CheckFileExt=False
		If FileExt="" or IsEmpty(FileExt) Then
			CheckFileExt = False
			Exit Function
		End If
		If FileExt="asp" or FileExt="asa" or FileExt="aspx" Then
			CheckFileExt = False
			Exit Function
		End If
		Forumupload = Split(InceptFile,",")
		For i = 0 To ubound(Forumupload)
			If FileExt = Trim(Forumupload(i)) Then
				CheckFileExt = True
				Exit Function
			Else
				CheckFileExt = False
			End If
		Next
	End Function

	'-----------------------------------------------------------------------------------
	'判断文件类型:0=其它,1=图片,2=FLASH,3=音乐,4=电影
	'-----------------------------------------------------------------------------------
	Private Function CheckFiletype(Byval FileExt)
		FileExt = Lcase(Replace(FileExt,".",""))
		Select Case FileExt
				Case "gif", "jpg", "jpeg","png","bmp","tif","iff"
					CheckFiletype=1
				Case "swf", "swi"
					CheckFiletype=2
				Case "mid", "wav", "mp3","rmi","cda"
					CheckFiletype=3
				Case "avi", "mpg", "mpeg","ra","ram","wov","asf"
					CheckFiletype=4
				Case Else
					CheckFiletype=0
		End Select
	End Function

	'-----------------------------------------------------------------------------------
	'执行保存上传文件
	'-----------------------------------------------------------------------------------
	Public Sub SaveUpFile()
		On Error Resume Next
		Select Case (Upload_Type) 
			Case 0
				ObjName = "无组件"
				
				Set UploadObj = New UpFile_Class
				If Err.Number<>0 Then
					ErrCodes = 1
				Else					
					SaveFile_0
				End If
			Case 1
				ObjName = "Aspupload3.0组件"
				Set UploadObj = Server.CreateObject("Persits.Upload") 
				If Err.Number<>0 Then
					ErrCodes = 1
				Else
					SaveFile_1
				End If
			Case 2
				ObjName = "SA-FileUp 4.0组件"
				Set UploadObj = Server.CreateObject("SoftArtisans.FileUp")
				If Err.Number<>0 Then
					ErrCodes = 1
				Else
					SaveFile_2
				End If
			Case 3
				ObjName = "DvFile.Upload V1.0组件"
				Set UploadObj = Server.CreateObject("DvFile.Upload")
				If Err.Number<>0 Then
					ErrCodes = 1
				Else
					SaveFile_3
				End If
			Case Else
				ErrCodes = 2
		End Select
	End Sub

	''-----------------------------------------------------------------------------------
	' 上传处理过程
	''-----------------------------------------------------------------------------------
	''-----------------------------------------------------------------------------------
	''无组件上传
	''-----------------------------------------------------------------------------------
	Private Sub SaveFile_0()
		Dim FormName,Item,File
		Dim FileExt,OldFileName,FileName,FileType,FileToBinary 'OldFileName By Clwang
		UploadObj.InceptFileType = InceptFile
		UploadObj.MaxSize = FileMaxSize
		UploadObj.GetDate ()	'取得上传数据
		FileToBinary = Null

		
		If Not IsEmpty(SessionName) Then			

			If Session(SessionName) <> UploadObj.Form(SessionName) or Session(SessionName) = Empty Then
				ErrCodes = 7
				Exit Sub
			End If
			
		End If
		If UploadObj.Err > 0 then
			Select Case UploadObj.Err
				Case 1 : ErrCodes = 3
				Case 2 : ErrCodes = 4
				Case 3 : ErrCodes = 5
			End Select
			Exit Sub
		Else
			For Each FormName In UploadObj.File		''列出所有上传了的文件
				If Count>MaxFile Then
					ErrCodes = 6
					Exit Sub
				End If
				Set File = UploadObj.File(FormName)
				FileExt = FixName(File.FileExt)
				If CheckFileExt(FileExt) = False then
					ErrCodes = 5
					EXIT SUB
				End If
				OldFileName=File.FileName
				FileName = FormatName(OldFileName,FileExt)
				FileType = CheckFiletype(FileExt)
				If IsBinary Then
					FileToBinary = File.FileData
				End If
				If File.FileSize>0 Then
					File.SaveToFile Server.Mappath(FilePath & FileName)
					AddData FormName , _ 
							OldFileName , _
							FileName , _
							FilePath , _
							File.FileSize , _
							File.FileType , _
							FileType , _
							FileToBinary , _
							FileExt , _
							File.FileWidth , _
							File.FileHeight
					Count = Count + 1
					CountSize = CountSize + File.FileSize
				End If
				Set File=Nothing
			Next
			For Each Item in UploadObj.Form
				If UploadForms.Exists (Item) Then _
					UploadForms(Item) = UploadForms(Item) & ", " & UploadObj.Form(Item) _
				Else _
				UploadForms.Add Item , UploadObj.Form(Item)
			Next
			If Not IsEmpty(SessionName) Then Session(SessionName) = Empty
		End If
	End Sub
	''-----------------------------------------------------------------------------------
	''Aspupload3.0组件上传
	''-----------------------------------------------------------------------------------
	Private Sub SaveFile_1()
		Dim FileCount
		Dim FormName,Item,File
		Dim FileExt,OldFileName,FileName,FileType,FileToBinary 'OldFileName By Clwang
		UploadObj.OverwriteFiles = False		'不能复盖
		UploadObj.IgnoreNoPost = True
		UploadObj.SetMaxSize FileMaxSize, True	'限制大小
		UploadObj.ProgressID = ProgressID_Str
		FileCount = UploadObj.Save
		FileToBinary = Null
		If Not IsEmpty(SessionName) Then
			If Session(SessionName) <> UploadObj.Form(SessionName) or Session(SessionName) = Empty Then
				ErrCodes = 7
				Exit Sub
			End If
		End If

		If Err.Number = 8 Then
				ErrCodes = 4
				EXIT SUB
		Else 
				If Err <> 0 Then
					ErrCodes = -1
					Response.Write "错误信息: " & Err.Description
					EXIT SUB
				End If
				If FileCount < 1 Then 
					ErrCodes = 3
					EXIT SUB
				End If
				For Each File In UploadObj.Files	'列出所有上传文件
					If Count>MaxFile Then
						ErrCodes = 6
						Exit Sub
					End If
					FileExt = FixName(Replace(File.Ext,".",""))
					If CheckFileExt(FileExt) = False then
						ErrCodes = 5
						EXIT SUB
					End If
					OldFileName=File.FileName
					FileName = FormatName(OldFileName,FileExt)
					FileType = CheckFiletype(FileExt)
					If IsBinary Then
						FileToBinary = File.Binary
					End If
					'File.Filename
					If File.Size>0 Then
						File.SaveAs Server.Mappath(FilePath & FileName)
						AddData File.Name , _ 
							OldFileName , _
							FileName , _
							FilePath , _
							File.Size , _
							File.ContentType , _
							FileType , _
							FileToBinary , _
							FileExt , _
							File.ImageWidth , _
							File.ImageHeight
						Count = Count + 1
						CountSize = CountSize + File.Size
					End If
				Next
				For Each Item in UploadObj.Form
					If UploadForms.Exists (Item) Then _
						UploadForms(Item) = UploadForms(Item) & ", " & Item.Value _
					Else _
						UploadForms.Add Item.Name , Item.Value
				Next
				If Not IsEmpty(SessionName) Then Session(SessionName) = Empty
		End If
	End Sub
	''-----------------------------------------------------------------------------------
	''SA-FileUp 4.0组件上传FileUpSE V4.09
	''-----------------------------------------------------------------------------------
	Private Sub SaveFile_2()
		Dim FormName,Item,File,FormNames
		Dim FileExt,OldFileName,FileName,FileType,FileToBinary 'OldFileName By Clwang
		Dim Filesize
		FileToBinary = Null
		If Not IsEmpty(SessionName) Then
			If Session(SessionName) <> UploadObj.Form(SessionName) or Session(SessionName) = Empty Then
				ErrCodes = 7
				Exit Sub
			End If
		End If
		For Each FormName In UploadObj.Form
			FormNames = ""
			If IsObject(UploadObj.Form(FormName)) Then
				If Not UploadObj.Form(FormName).IsEmpty Then
					UploadObj.Form(FormName).Maxbytes = FileMaxSize	'限制大小
					UploadObj.OverWriteFiles = False
					Filesize = UploadObj.Form(FormName).TotalBytes
					If Err.Number<>0 Then
						ErrCodes = -1
						Response.Write "错误信息: " & Err.Description
						EXIT SUB
					End If
					If Filesize>FileMaxSize then
						ErrCodes = 4
						Exit sub
					End If
					FileName	= UploadObj.Form(FormName).ShortFileName	 '原文件名
					FileExt		= Mid(Filename, InStrRev(Filename, ".")+1)
					FileExt		= FixName(FileExt)
					If CheckFileExt(FileExt) = False then
						ErrCodes = 5
						EXIT SUB
					End If
					OldFileName=FileName
					FileName = FormatName(OldFileName,FileExt)
					FileType = CheckFiletype(FileExt)
					'If IsBinary Then
						'FileToBinary = UploadContents (2)
					'End If
					'保存文件
					If Filesize>0 Then
						UploadObj.Form(FormName).SaveAs Server.MapPath(FilePath & FileName)
						AddData FormName , _ 
								OldFileName , _
								FileName , _
								FilePath , _
								FileSize , _
								UploadObj.Form(FormName).ContentType , _
								FileType , _
								FileToBinary , _
								FileExt , _
								-1 , _
								-1
						Count = Count + 1
						CountSize = CountSize + Filesize
					End If
				Else
					ErrCodes = 3
					EXIT SUB
				End If
			Else
				If UploadObj.FormEx(FormName).Count > 1 Then
					For Each FormNames In UploadObj.FormEx(FormName)
						FormNames = FormNames & ", " & FormNames
					Next
					UploadForms.Add FormName , FormNames
				Else
					UploadForms.Add FormName , UploadObj.Form(FormName)
				End If
			End If
		Next
		If Not IsEmpty(SessionName) Then Session(SessionName) = Empty
	End Sub
	''-----------------------------------------------------------------------------------
	''DvFile.Upload V1.0组件上传
	''-----------------------------------------------------------------------------------
	Private Sub SaveFile_3()
		Dim FormName,Item,File
		Dim FileExt,OldFileName,FileName,FileType,FileToBinary 'OldFileName By Clwang
		UploadObj.InceptFileType = InceptFile
		UploadObj.MaxSize = FileMaxSize
		UploadObj.Install
		FileToBinary = Null
		If Not IsEmpty(SessionName) Then
			If Session(SessionName) <> UploadObj.Form(SessionName) or Session(SessionName) = Empty Then
				ErrCodes = 7
				Exit Sub
			End If
		End If
		If UploadObj.Err > 0 then
			Select Case UploadObj.Err
				Case 1 : ErrCodes = 3
				Case 2 : ErrCodes = 4
				Case 3 : ErrCodes = 5
				Case 4 : ErrCodes = 5
				Case 5 : ErrCodes = -1
			End Select
			Exit Sub
		Else
			For Each FormName In UploadObj.File		''列出所有上传了的文件
				If Count>MaxFile Then
					ErrCodes = 6
					Exit Sub
				End If
				Set File = UploadObj.File(FormName)
				FileExt = FixName(File.FileExt)
				If CheckFileExt(FileExt) = False then
					ErrCodes = 5
					EXIT SUB
				End If
				OldFileName = File.FileName
				FileName = FormatName(OldFileName,FileExt)
				FileType = CheckFiletype(FileExt)
				If IsBinary Then
					FileToBinary = File.FileData
				End If
				If File.FileSize>0 Then
					UploadObj.SaveToFile Server.mappath(FilePath & FileName),FormName
					AddData FormName , _ 
							OldFileName , _
							FileName , _
							FilePath , _
							File.FileSize , _
							File.FileType , _
							FileType , _
							FileToBinary , _
							FileExt , _
							File.FileWidth , _
							File.FileHeight
					Count = Count + 1
					CountSize = CountSize + File.FileSize
				End If
				Set File=Nothing
			Next
			For Each Item in UploadObj.Form
				UploadForms.Add Item.Name , Item.Value
			Next
			If Not IsEmpty(SessionName) Then Session(SessionName) = Empty
		End If
	End Sub

	Private Sub AddData( Form_Name,OldFile_Name,File_Name,File_Path,File_Size,File_ContentType,File_Type,File_Data,File_Ext,File_Width,File_Height )
		Set FileInfo = New FileInfo_Cls
			FileInfo.FormName = Form_Name
			FileInfo.OldFileName = OldFile_Name  '原文件名，By梅傲风
			FileInfo.FileName = File_Name
			FileInfo.FilePath = File_Path
			FileInfo.FileSize = File_Size
			FileInfo.FileType = File_Type
			FileInfo.FileContentType = File_ContentType
			FileInfo.FileExt = File_Ext
			FileInfo.FileData = File_Data
			FileInfo.FileHeight = File_Height
			FileInfo.FileWidth = File_Width
			UploadFiles.Add Form_Name , FileInfo
		Set FileInfo = Nothing
	End Sub

	'创建预览图片:Call CreateView(原始文件的路径,预览文件名及路径,原文件后缀)
	Public Sub CreateView(Imagename,TempFilename,FileExt)
	'===========================================================
		If ErrCodes <>0 Then Exit Sub
		On Error Resume Next
		Select Case Preview_Type
		Case 0
		'---------------------CreatePreviewImage---------------
			ObjName = "CreatePreviewImage组件"
			Set ImageObj = Server.CreateObject("CreatePreviewImage.cGvbox")
		Case 1
		'---------------------AspJpegV1.2---------------
			ObjName = "AspJpegV1.2组件"
			Set ImageObj = Server.CreateObject("Persits.Jpeg")
		Case 2
		'---------------------SoftArtisans ImgWriter V1.21---------------
			ObjName = "SoftArtisans ImgWriter V1.21组件"
			Set ImageObj = Server.CreateObject("SoftArtisans.ImageGen")
		Case Else
			Preview_Type = 999 : Exit Sub
		End Select
		If Err.Number<>0 Then ErrCodes = 1 : Exit Sub
	'===========================================================
		Select Case Preview_Type
			Case 0
				Image_Obj_0 Imagename,TempFilename,FileExt
			Case 1
				Image_Obj_1 Imagename,TempFilename,FileExt
			Case 2
				Image_Obj_2 Imagename,TempFilename,FileExt
			Case Else
				Preview_Type = 999
		End Select
	End Sub

	Sub Image_Obj_0(Imagename,TempFilename,FileExt)
			ImageObj.SetSavePreviewImagePath = Server.MapPath(TempFilename)			'预览图存放路径
			ImageObj.SetPreviewImageSize = SetPreviewImageSize						'预览图宽度
			ImageObj.SetImageFile = Trim(Server.MapPath(Imagename))					'Imagename原始文件的物理路径
			'创建预览图的文件
			If ImageObj.DoImageProcess = False Then
				ErrCodes = -1
				Response.Write "生成预览图错误: " & ImageObj.GetErrString
			End If
	End Sub

	'---------------------AspJpegV1.2---------------
	Sub Image_Obj_1(Imagename,TempFilename,FileExt)
			' 读取要处理的原文件
			Dim Draw_X,Draw_Y,Logobox
			Draw_X = 0
			Draw_Y = 0
			FileExt = Lcase(FileExt)
			ImageObj.Open Trim(Server.MapPath(Imagename))
			If ImageObj.OriginalWidth<View_ImageWidth or ImageObj.Originalheight<View_ImageHeight Then
				TempFilename = ""
				Exit Sub
			Else
				If FileExt<>"gif" and ImageObj.OriginalWidth > Draw_ImageWidth * 2 and Draw_Type >0 Then
					Draw_X = DrawImage_X(ImageObj.OriginalWidth,Draw_ImageWidth,2)
					Draw_Y = DrawImage_y(ImageObj.Originalheight,Draw_ImageHeight,2)
					If Draw_Type=2 Then
						Set Logobox = Server.CreateObject("Persits.Jpeg")
						'*添加水印图片	添加时请关闭水印字体*
						'//读取添加的图片
						Logobox.Open Server.MapPath(Draw_Info)
						Logobox.Width = Draw_ImageWidth								'// 加入图片的原宽度
						Logobox.Height = Draw_ImageHeight							'// 加入图片的原高度
						ImageObj.DrawImage Draw_X, Draw_Y, Logobox, Draw_Graph,Transition_Color,90	'// 加入图片的位置价坐标（添加水印图片）
						'ImageObj.Sharpen 1, 130
						ImageObj.Save Server.MapPath(Imagename)
						Set Logobox=Nothing
					Else
						'//关于修改字体及文字颜色的
						ImageObj.Canvas.Font.Color		= Draw_FontColor	'// 文字的颜色
						ImageObj.Canvas.Font.Family		= Draw_FontFamily	'// 文字的字体
						ImageObj.Canvas.Font.Bold		= Draw_FontBold
						ImageObj.Canvas.Font.Size		= Draw_FontSize					'//字体大小
						' Draw frame: black, 2-pixel width
						ImageObj.Canvas.Print Draw_X, Draw_Y, Draw_Info	'// 加入文字的位置坐标
						ImageObj.Canvas.Pen.Color		= &H000000		'// 边框的颜色
						ImageObj.Canvas.Pen.Width		= 1				'// 边框的粗细
						ImageObj.Canvas.Brush.Solid	= False			'// 图片边框内是否填充颜色
						'ImageObj.Canvas.Bar 0, 0, ImageObj.Width, ImageObj.Height	'// 图片边框线的位置坐标
						ImageObj.Save Server.MapPath(Imagename)
					End If
				End If
				If ImageObj.Width > ImageObj.height Then
					ImageObj.Width = View_ImageWidth
					ImageObj.Height = ViewImage_Height(ImageObj.OriginalWidth,ImageObj.Originalheight,View_ImageWidth,View_ImageHeight)
				Else
					ImageObj.Width = ViewImage_Width(ImageObj.OriginalWidth,ImageObj.Originalheight,View_ImageWidth,View_ImageHeight)
					ImageObj.Height = View_ImageHeight
				End If
				ImageObj.Sharpen 1, 120
				ImageObj.Save Server.MapPath(TempFilename)		'// 生成预览文件
			End If
	End Sub

	'SoftArtisans ImgWriter V1.21
	Public Sub Image_Obj_2(Imagename,TempFilename,FileExt)
			'定义变量
			Dim Draw_X,Draw_Y
			FileExt = Lcase(FileExt)
			Draw_X = 0
			Draw_Y = 0
			' 读取要处理的原文件
			ImageObj.LoadImage Trim(Server.MapPath(Imagename))
			If ImageObj.ErrorDescription <> "" Then
				TempFilename = ""
				ErrCodes = -1
				Response.Write "生成预览图错误: " &ImageObj.ErrorDescription
				Exit Sub
			End If
			If ImageObj.Width<Cint(View_ImageWidth) or ImageObj.Height<Cint(View_ImageHeight) Then
				TempFilename=""
				Exit Sub
			Else
				IF FileExt<>"gif" and ImageObj.Width > Draw_ImageWidth * 2 and Draw_Type>0 Then
					Draw_X = DrawImage_X(ImageObj.Width,Draw_ImageWidth,2)
					Draw_Y = DrawImage_y(ImageObj.Height,Draw_ImageHeight,2)
					Dim saiTopMiddle
					Select Case Draw_XYType
						Case "0" '左上
							saiTopMiddle = 3
						Case "1" '左下
							saiTopMiddle = 5
						Case "2" '居中
							saiTopMiddle = 1
						Case "3" '右上
							saiTopMiddle = 6
						Case "4" '右下
							saiTopMiddle = 8
						Case Else '不显示
							saiTopMiddle = 0
					End Select
					If Draw_Type=2 Then
						ImageObj.AddWatermark Server.MapPath(Draw_Info), saiTopMiddle, Draw_Graph,Transition_Color,True
						'ImageObj.AddWatermark Server.MapPath(Request.QueryString("mimg")), 0, 0.3
					Else
						ImageObj.Font.Italic	= False			'斜体
						ImageObj.Font.height	= Draw_FontSize
						ImageObj.Font.name		= Draw_FontFamily
						ImageObj.Font.Color		= Draw_FontColor
						ImageObj.Text			= Draw_Info
						ImageObj.DrawTextOnImage Draw_X, Draw_Y, ImageObj.TextWidth, ImageObj.TextHeight
					End If
					ImageObj.SaveImage 0, ImageObj.ImageFormat, Server.MapPath(Imagename)
				End If
				'ImageObj.SharpenImage 100
				ImageObj.ColorResolution = 24	'24色保存
				ImageObj.ResizeImage View_ImageWidth,View_ImageHeight,0,0
				'0=saiFile,1=saiMemory,2=saiBrowser,4=saiDatabaseBlob
				'saiBMP=1,saiGIF=2,saiJPG=3,saiPNG=4,saiPCX=5,saiTIFF=6,saiWMF=7,saiEMF=8,saiPSD=9 
				ImageObj.SaveImage 0, 3, Server.MapPath(TempFilename)
			End If
	End Sub

	'比例或固定缩小
	Private Function ViewImage_Width(Image_W,Image_H,xView_W,xView_H)
		If Draw_SizeType = "1" Then
			ViewImage_Width = Image_W * xView_H / Image_H
		Else
			ViewImage_Width = xView_W
		End If
	End Function

	Private Function ViewImage_Height(Image_W,Image_H,xView_W,xView_H)
		If Draw_SizeType = "1" Then
			ViewImage_Height = xView_W * Image_H / Image_W
		Else
			ViewImage_Height = xView_H
		End If
	End Function

	'SpaceVal X轴坐标边缘距离
	Private Function DrawImage_X(xImage_W,xLogo_W,SpaceVal)
		Select Case Draw_XYType
			Case "0" '左上
				DrawImage_X = SpaceVal
			Case "1" '左下
				DrawImage_X = SpaceVal
			Case "2" '居中
				DrawImage_X = (xImage_W + xLogo_W) / 2
			Case "3" '右上
				DrawImage_X = xImage_W - xLogo_W - SpaceVal
			Case "4" '右下
				DrawImage_X = xImage_W - xLogo_W - SpaceVal
			Case Else '不显示
				DrawImage_X = 0
		End Select
	End Function

	'SpaceVal Y轴坐标边缘距离
	Private Function DrawImage_Y(yImage_H,yLogo_H,SpaceVal)
		Select Case Draw_XYType
			Case "0" '左上
				DrawImage_Y = SpaceVal
			Case "1" '左下
				DrawImage_Y = yImage_H - yLogo_H - SpaceVal
			Case "2" '居中
				DrawImage_Y = (yImage_H + yLogo_H) / 2
			Case "3" '右上
				DrawImage_Y = SpaceVal
			Case "4" '右下
				DrawImage_Y = yImage_H - yLogo_H - SpaceVal
			Case Else '不显示
				DrawImage_Y = 0
		End Select
	End Function

End Class

Class FileInfo_Cls
	Public FormName,OldFileName,FileName,FilePath,FileSize,FileContentType,FileType,FileData,FileExt,FileWidth,FileHeight
	Private Sub Class_Initialize
		FileWidth = -1
		FileHeight = -1
	End Sub
End Class


Class UpFile_Class
	Public Form,File,Version,Err
	Private CHK_FileType,CHK_MaxSize

	Private Sub Class_Initialize
		Version = "无惧上传类 Version V1.0"
		Err = -1
		CHK_FileType = ""
		CHK_MaxSize = -1
		Set Form = Server.CreateObject ("Scripting.Dictionary")
		Set File = Server.CreateObject ("Scripting.Dictionary")
		Set oUpFileStream = Server.CreateObject ("Adodb." & "Str" & "eam")
		Form.CompareMode = 1
		File.CompareMode = 1
		oUpFileStream.Type = 1
		oUpFileStream.Mode = 3
		oUpFileStream.Open
	End Sub

	Private Sub Class_Terminate  
		'清除变量及对像
		Form.RemoveAll
		Set Form = Nothing
		File.RemoveAll
		Set File = Nothing
		oUpFileStream.Close
		Set oUpFileStream = Nothing
	End Sub

	Public Property Get InceptFileType
		InceptFileType = CHK_FileType
	End Property
	Public Property Let InceptFileType(Byval vType)
		CHK_FileType = vType
	End Property

	Public Property Get MaxSize
		MaxSize = CHK_MaxSize
	End Property
	Public Property Let MaxSize(vSize)
		If IsNumeric(vSize) Then CHK_MaxSize = Int(vSize)
	End Property

	Public Sub GetDate()
	   '定义变量
	  Dim RequestBinDate,sSpace,bCrLf,sInfo,iInfoEnd,tStream,iStart,oFileInfo
	  Dim sFormValue,sFileName,sFormName,RequestSize
	  Dim iFindStart,iFindEnd,iFormStart,iFormEnd,FileBlag
	   '代码开始
	  RequestSize = Request.TotalBytes
	  if Not IsNumeric(RequestSize) then RequestSize=0
	  RequestSize = Int(RequestSize)
	  If RequestSize < 1 Then
		Err = 1
		Exit Sub
	  End If

	  Set tStream = Server.CreateObject ("Adodb." & "Str" & "eam")
	  oUpFileStream.Write Request.BinaryRead (RequestSize)
	  oUpFileStream.Position = 0
	  RequestBinDate = oUpFileStream.Read
	  iFormEnd = oUpFileStream.Size
	  
	  bCrLf = ChrB (13) & ChrB (10)
	  '取得每个项目之间的分隔符
	  sSpace = MidB (RequestBinDate,1, InStrB (1,RequestBinDate,bCrLf)-1)
	  iStart = LenB  (sSpace)
	  iFormStart = iStart+2
	  '分解项目
	  Do
	    iInfoEnd = InStrB (iFormStart,RequestBinDate,bCrLf & bCrLf)+3
	    tStream.Type = 1
	    tStream.Mode = 3
	    tStream.Open
	    oUpFileStream.Position = iFormStart
	    oUpFileStream.CopyTo tStream,iInfoEnd-iFormStart
	    tStream.Position = 0
	    tStream.Type = 2
	    tStream.CharSet = "gb2312"
	    sInfo = tStream.ReadText
	    '取得表单项目名称
	    iFormStart = InStrB (iInfoEnd,RequestBinDate,sSpace)-1
	    iFindStart = InStr(22,sInfo,"name=""",1)+6
	    iFindEnd = InStr(iFindStart,sInfo,"""",1)
	    sFormName = Mid(sinfo,iFindStart,iFindEnd-iFindStart)
	    '如果是文件
		If InStr(45,sInfo,"filename=""",1) > 0 Then
			Set oFileInfo = new FileInfo_Class
			'取得文件属性
			iFindStart = InStr(iFindEnd,sInfo,"filename=""",1)+10
			iFindEnd = InStr(iFindStart,sInfo,"""",1)
			sFileName = Mid(sinfo,iFindStart,iFindEnd-iFindStart)
			oFileInfo.FileName = Mid(sFileName,InStrRev(sFileName, "\")+1)
			oFileInfo.FilePath = Left(sFileName,InStrRev(sFileName, "\"))
			oFileInfo.FileExt = Lcase(Mid(sFileName,InStrRev(sFileName, ".")+1))
			iFindStart = InStr (iFindEnd,sInfo,"Content-Type: ",1)+14
			iFindEnd = InStr (iFindStart,sInfo,vbCr)
			oFileInfo.FileType = Ucase(Mid(sinfo,iFindStart,iFindEnd-iFindStart))
			oFileInfo.FileStart = iInfoEnd
			oFileInfo.FileSize = iFormStart -iInfoEnd -2
			oFileInfo.FormName = sFormName
			If Instr(oFileInfo.FileType,"IMAGE/") Or Instr(oFileInfo.FileType,"FLASH") Then
				FileBlag = GetImageSize
				oFileInfo.FileExt = FileBlag(0)
				oFileInfo.FileWidth = FileBlag(1)
				oFileInfo.FileHeight = FileBlag(2)
				FileBlag = Empty
			End If
			If CHK_MaxSize > 0 Then
				If oFileInfo.FileSize > CHK_MaxSize Then
					Err = 2
					Exit Sub
				End If
			End If
			If CheckErr(oFileInfo.FileExt) = False Then
				Exit Sub
			End If
			File.Add sFormName,oFileInfo
		Else
			'如果是表单项目
			tStream.Close
			tStream.Type = 1
			tStream.Mode = 3
			tStream.Open
			oUpFileStream.Position = iInfoEnd 
			oUpFileStream.CopyTo tStream,iFormStart-iInfoEnd-2
			tStream.Position = 0
			tStream.Type = 2
			tStream.CharSet = "gb2312"
			sFormValue = tStream.ReadText
			If Form.Exists (sFormName) Then _
				Form (sFormName) = Form (sFormName) & ", " & sFormValue _
			Else _
				Form.Add sFormName,sFormValue
		End If
		tStream.Close
		iFormStart = iFormStart+iStart+2
	  '如果到文件尾了就退出
	  Loop Until  (iFormStart+2) = iFormEnd
	  RequestBinDate = ""
	  Set tStream = Nothing
	End Sub

	'====================================================================
	'验证上传类型
	'====================================================================
	Private Function CheckErr(Byval ChkExt)
		CheckErr=False
		If CHK_FileType = "" Then CheckErr=True : Exit Function
		Dim ChkStr
		ChkStr = ","&Lcase(CHK_FileType)&","
		If Instr(ChkStr,","&ChkExt&",")>0 Then
			CheckErr=True
		Else
			Err = 3
		End If
	End Function
	'====================================================================
	'图像宽高类型读取
	'====================================================================
	Private Function Bin2Str(Byval Bin)
		Dim i, Str, Sclow
		For i = 1 To LenB(Bin)
			Sclow = MidB(Bin,i,1)
			If ASCB(Sclow)<128 Then
				Str = Str & Chr(ASCB(Sclow))
			Else
				i = i+1
				If i <= LenB(Bin) Then Str = Str & Chr(ASCW(MidB(Bin,i,1)&Sclow))
			End If
		Next 
		Bin2Str = Str
	End Function

	Private Function Num2Str(Byval num,Byval Base,Byval Lens)
		Dim ImageSize
		ImageSize = ""
		While(num>=Base)
			ImageSize = (num mod Base) & ImageSize
			num = (num - num mod Base)/Base
		Wend
		Num2Str = Right(String(Lens,"0") & num & ImageSize,Lens)
	End Function

	Private Function Str2Num(Byval str,Byval Base)
		Dim ImageSize,i
		ImageSize = 0
		For i=1 To Len(str)
			ImageSize = ImageSize *Base + Cint(Mid(str,i,1))
		Next
		Str2Num = ImageSize
	End Function

	Private Function BinVal(Byval bin)
		Dim ImageSize,i
		ImageSize = 0
		For i = lenb(bin) To 1 Step -1
			ImageSize = ImageSize *256 + ASCB(Midb(bin,i,1))
		Next
		BinVal = ImageSize
	End Function

	Private Function BinVal2(Byval bin)
		Dim ImageSize,i
		ImageSize = 0
		For i = 1 To Lenb(bin)
			ImageSize = ImageSize *256 + ASCB(Midb(bin,i,1))
		Next
		BinVal2 = ImageSize
	End Function

	Private Function GetImageSize() 
		Dim ImageSize(2),bFlag
		bFlag = oUpFileStream.Read(3)

		Select Case Hex(BinVal(bFlag))
			Case "4E5089":
				oUpFileStream.Read(15)
				ImageSize(0) = "png"
				ImageSize(1) = BinVal2(oUpFileStream.Read(2))
				oUpFileStream.Read(2)
				ImageSize(2) = BinVal2(oUpFileStream.Read(2))
			Case "464947":
				oUpFileStream.Read(3)
				ImageSize(0) = "gif"
				ImageSize(1) = BinVal(oUpFileStream.Read(2))
				ImageSize(2) = BinVal(oUpFileStream.Read(2))
			Case "535746":
				Dim BinData,sConv,nBits
				oUpFileStream.Read(5)
				BinData = oUpFileStream.Read(1)
				sConv = Num2Str(ASCB(BinData),2 ,8)
				nBits = Str2Num(Left(sConv,5),2)
				sConv = Mid(sConv,6)
				While(Len(sConv)<nBits*4)
					BinData = oUpFileStream.Read(1)
					sConv = sConv&Num2Str(ASCB(BinData),2 ,8)
				Wend
				ImageSize(0) = "swf"
				ImageSize(1) = Int(ABS(Str2Num(Mid(sConv,1*nBits+1,nBits),2)-Str2Num(Mid(sConv,0*nBits+1,nBits),2))/20)
				ImageSize(2) = Int(ABS(Str2Num(Mid(sConv,3*nBits+1,nBits),2)-Str2Num(Mid(sConv,2*nBits+1,nBits),2))/20)
			Case "535743":'flashmx
				ImageSize(0) = "swf"
				ImageSize(1) = 0
				ImageSize(2) = 0
			Case "FFD8FF":
				Dim p1
				Do 
					Do: p1 = BinVal(oUpFileStream.Read(1)): Loop While p1 = 255 And Not oUpFileStream.EOS
					If p1>191 and p1<196 Then Exit Do Else oUpFileStream.Read(BinVal2(oUpFileStream.Read(2))-2)
					Do:p1 = BinVal(oUpFileStream.Read(1)):Loop While p1<255 And Not oUpFileStream.EOS
					Loop While True
					oUpFileStream.Read(3)
					ImageSize(0) = "jpg"
					ImageSize(2) = BinVal2(oUpFileStream.Read(2))
					ImageSize(1) = BinVal2(oUpFileStream.Read(2))
			Case Else:
				If Left(Bin2Str(bFlag),2) = "BM" Then
					oUpFileStream.Read(15)
					ImageSize(0) = "bmp"
					ImageSize(1) = BinVal(oUpFileStream.Read(4))
					ImageSize(2) = BinVal(oUpFileStream.Read(4))
				Else
					ImageSize(0) = "(UNKNOWN)"
				End If
		End Select
		GetImagesize = ImageSize
	End Function
End Class

'文件属性类
Class FileInfo_Class
	Public FormName,FileName,FilePath,FileSize,FileType,FileStart,FileExt,FileWidth,FileHeight
	Private Sub Class_Initialize
		FileWidth=0
		FileHeight=0
	End Sub
	'保存文件方法
	Public Sub SaveToFile (Byval Path)
		Dim Ext,oFileStream
		Ext = LCase(Mid(Path, InStrRev(Path, ".") + 1))
		If Ext <> FileExt Then Exit Sub
		If Trim(Path)="" or FileStart=0 or FileName="" or Right(Path,1)="/" Then Exit Sub
		On Error Resume Next
		Set oFileStream = CreateObject ("Adodb." & "Str" & "eam")
		oFileStream.Type = 1
		oFileStream.Mode = 3
		oFileStream.Open
		oUpFileStream.Position = FileStart
		oUpFileStream.CopyTo oFileStream,FileSize
		oFileStream.SaveToFile Path,2
		oFileStream.Close
		Set oFileStream = Nothing 
	End Sub
	'取得文件数据
	Public Function FileData
		oUpFileStream.Position = FileStart
		FileData = oUpFileStream.Read (FileSize)
	End Function
End Class