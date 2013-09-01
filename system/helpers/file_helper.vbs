Function GetData(byref url, byref GetMode)

End Function

'================================================
' 函数名：ChkMapPath
' 作  用：相对路径转换为绝对路径
' 参  数：strPath ----原路径
' 返回值：绝对路径
'================================================
Public Function ChkMapPath(ByVal strPath)
	On Error Resume Next
	Dim fullPath
	strPath = Replace(Replace(Trim(strPath), "//", "/"), "\\", "\")

	If strPath = "" Then strPath = "."
	If InStr(strPath,":\") = 0 Then 
		fullPath = Server.MapPath(strPath)
	Else
		strPath = Replace(strPath,"/","\")
		fullPath = Trim(strPath)
		If Right(fullPath, 1) = "\" Then
			fullPath = Left(fullPath, Len(fullPath) - 1)
		End If
	End If
	ChkMapPath = fullPath
End Function


' ***********************************
'   删除文件
' ***********************************
Function DeleteFiles(filepath)

    Dim FSO
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")

    If FSO.FileExists(filepath) Then
        FSO.DeleteFile filepath, True
        DeleteFiles = True
    Else
        DeleteFiles = False
    End If

    Set FSO = Nothing

End Function

' ***********************************
'   判断文件是否存在
'	@filepath			相对路径(test\test.htm)
' ***********************************
Function FileExist(filepath)

  FileExist = False
  Dim FSO
  Set FSO = Server.CreateObject("Scripting.FileSystemObject")
  'filepath = Server.MapPath(".") & "\" & filepath

  If FSO.FileExists(filepath) Then
     FileExist = True
  End If	
End Function

' ***********************************
'	复制文件
' ***********************************
Function CopyFiles(tempSource, tempEnd)

    Dim CopyFSO
    Set CopyFSO = Server.CreateObject("Scripting.FileSystemObject")
    
    If CopyFSO.FileExists(tempEnd) Then
       CopyFiles = "目标备份文件 <b>" & tempEnd & "</b> 已存在，请先删除!"
       Set CopyFSO = Nothing
       Exit Function
    End If
    
    If CopyFSO.FileExists(tempSource) Then
    Else
       CopyFiles = "要复制的源数据库文件 <b>" & tempSource & "</b> 不存在!"
       Set CopyFSO = Nothing
       Exit Function
    End If
    
    CopyFSO.CopyFile tempSource, tempEnd
    CopyFiles = "已经成功复制文件 <b>" & tempSource & "</b> 到 <b>" & tempEnd & "</b>"
    Set CopyFSO = Nothing
    
End Function

' ***********************************
'	获取文件信息
' ***********************************
Function GetFileInfo(filename)

    Dim FSO, File, FileInfo(3)
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
    
    If FSO.FileExists(filename) Then
    
        Set File = FSO.GetFile(filename)        
        FileInfo(0) = File.Size

        If FileInfo(0) / 1000 > 1 Then
            FileInfo(0) = Int(FileInfo(0) / 1000) & " KB"
        Else
            FileInfo(0) = FileInfo(0) & " Bytes"
        End If
    
        FileInfo(1) = LCase(Right(filename, 4))
        FileInfo(2) = File.DateCreated
        FileInfo(3) = File.Type
    
    End If
    
    GetFileInfo = FileInfo
    Set FSO = Nothing
    
End Function

Function GetFileSize(filename)
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	Set File = fso.GetFile(filename)        
    GetFileSize = File.Size
	Set fso = Nothing
End Function

' ***********************************
'	获取文件大小
' ***********************************
Function GetTotalSize(GetLocal, GetType)

    Dim FSO
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
    
    If Err <> 0 Then
        Err.Clear
        GetTotalSize = "Fail"
    Else
        Dim SiteFolder
        If GetType = "Folder" Then
            Set SiteFolder = FSO.GetFolder(GetLocal)
        Else
            Set SiteFolder = FSO.GetFile(GetLocal)
        End If
        GetTotalSize = SiteFolder.Size

        If GetTotalSize > 1024 * 1024 Then
			GetTotalSize = GetTotalSize / 1024 / 1024
		End If

        If InStr(GetTotalSize, ".") Then
            GetTotalSize = Left(GetTotalSize, InStr(GetTotalSize, ".") + 2)
            GetTotalSize = GetTotalSize & " MB"
        Else
            GetTotalSize = Fix(GetTotalSize / 1024) & " KB"
        End If
        
        Set SiteFolder = Nothing
    End If
    
    Set FSO = Nothing
    
End Function

' ***********************************
'	获取文件大小
' ***********************************
Public Function ShowSize(size)

    On Error Resume Next

    If size = "" Or IsNull(size) Then
        ShowSize = "0Byte"
        Exit Function
    End If

    ShowSize = size & "Byte"

    If size < 0 Then
        ShowSize = "0KB"
        Exit Function
    End If

    If size > 1024 Then
       size = (size \ 1024)
       ShowSize = size & "KB"
    End If

    If size > 1024 Then
       size = (size / 1024)
       ShowSize = FormatNumber(size, 2) & "MB"
    End If

    If size > 1024 Then
       size = (size / 1024)
       ShowSize = FormatNumber(size, 2) & "GB"
    End If

End Function

' ***********************************
'	读取文件, gb2312读取文件
' ***********************************
Public Function ReadFile(filepath)

    On Error Resume Next

    Dim fs, file

    Set fs = Server.CreateObject("Scripting.FileSystemObject") 

    Set file = fs.OpenTextFile(filepath, 1, True)
    ReadFile = file.ReadAll
    Set fs = Nothing
    Set file = Nothing

End Function


' 读取整个文本文件
Function ReadTextFile(filepath, charset)
    dim retval
    set stm=server.CreateObject("adodb.stream")
    stm.Type = 2 '以本模式读取
    stm.mode = 3 
    stm.charset = charset
    stm.open
    stm.loadfromfile filepath
    retval =stm.readtext
    stm.Close
    Set stm = Nothing
    ReadTextFile = retval
End Function

' 写整个文本文件
Sub WriteTextFile(filepath, byval Str, CharSet) 
    set stm=server.CreateObject("adodb.stream")
    stm.Type = 2 '以本模式读取
    stm.mode =3
    stm.charset = CharSet
    stm.open
    stm.WriteText str
    stm.SaveToFile filepath, 2 
    stm.flush
    stm.Close
    set stm=nothing
End Sub


' ***********************************
'	写文件
' ***********************************
Public Function BuildFile(ByVal file, ByVal content, is_gb2312)
    Dim fso, filestream
    If is_gb2312 = 1 Then
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
        'Response.Write "目录1：" & sFile & "<br>"
        Set filestream = fso.CreateTextFile(file, True)
        filestream.Write content
        filestream.Close
        Set filestream = Nothing
        Set fso = Nothing
    Else
        Set filestream = Server.CreateObject("ADODB.Stream")
        With filestream
            .Type = 2
            .Mode = 3
            .Open
            .Charset = "utf-8"
            '.Charset = "gb2312"
            .Position = filestream.size
            .WriteText = content
            .SaveToFile file, 2
            .Close
        End With
        Set filestream = Nothing
    End If
End Function

' ***********************************
'	检查文件夹是否存在
' ***********************************
Function CheckDir2(ByVal folderpath)
	Dim fso
	folderpath = Server.MapPath(".") & "\" & folderpath
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FolderExists(FolderPath) then
		'存在
		CheckDir2 = True
	Else
		'不存在
		CheckDir2 = False
	End if
	Set fso = nothing
End Function

' ***********************************
'	创建新的文件夹
' ***********************************
Function MakeDir2(ByVal foldername)
	Dim fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	fso.CreateFolder(Server.MapPath(".") &"\" & foldername)
	If fso.FolderExists(Server.MapPath(".") &"\" &foldername) Then
		MakeDir2 = True
	Else
		MakeDir2 = False
	End If
	Set fso = nothing
End Function

' ***********************************
'	创建文件夹
' ***********************************
Function CreateDirectory(directoryName)

    Dim fso
    Dim f
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(directoryName) Then
        f = fso.CreateFolder(directoryName)		
    End If
    
	Set f = Nothing
	Set fso = Nothing
End Function

Function GetDirectoryName(filepath)
	GetDirectoryName = Mid(filepath, 1, InStrRev(filepath, "\"))
End Function

' ***********************************
' * 创建目录
' * @strPath		目录串
' ***********************************
Public Function CreateFolder(strPath)
	strPath = Replace(strPath,"/", "\")
	arrPath = Split(strPath, "\")
	Dim Fso, I, tmpPath, arrPath
	Set Fso = Server.CreateObject("Scripting.FileSystemObject")
	tmpPath = Server.MapPath(".") 
	For I=0 To UBound(arrPath)
		tmpPath  = tmpPath & "\" & arrPath(I)
		If Not Fso.FolderExists(tmpPath) Then
			Fso.CreateFolder tmpPath
		End If
	Next
End Function

Function MoveFile(source, destination)
   Dim fso
   Set fso = CreateObject("Scripting.FileSystemObject")
   fso.MoveFile source, destination
   Set fso = Nothing
End Function

' ***********************************
'	获取文件扩展名
' ***********************************
Public Function GetFileEx(fileName)
    GetFileEx = Mid(fileName, InStrRev(fileName, ".") + 1)
End Function


' *********************************
'	字节转为BSTR
' *********************************
Function BytesToBstr(body, Cset)
    Dim objstream
    Set objstream = Server.CreateObject("ADO" & "DB.St" & "ream")
    objstream.Type = 1
    objstream.Mode = 3
    objstream.Open
    objstream.Write body
    objstream.Position = 0
    objstream.Type = 2
    objstream.Charset = Cset
    BytesToBstr = objstream.ReadText
    objstream.Close
    Set objstream = Nothing
End Function

' *********************************
'	读取远程xml
' *********************************
Public Function ReadRemoteFile(url, encode)
    Dim objXML
    Set objXML = server.CreateObject("MSXML2.XMLHTTP") '定义
    objXML.open "GET", url, False '打开
    objXML.send '发送
    If objXML.readystate <> 4 Then '判断文档是否已经解析完，以做客户端接受返回消息
        Exit Function
    End If
        
        
    ReadRemoteFile = BytesToBstr(objXML.responseBody, encode)
    
    Set objXML = Nothing '关闭
    If Err.Number <> 0 Then Err.Clear
End Function

Function GetContent(str,start,last,n)

	If Instr(lcase(str),lcase(start))>0 and Instr(lcase(str),lcase(last))>0 then
		select case n
		case 0	'左右都截取（都取前面）（去处关键字）
		GetContent=Right(str,Len(str)-Instr(lcase(str),lcase(start))-Len(start)+1) 
		GetContent=Left(GetContent,Instr(lcase(GetContent),lcase(last))-1)
		case 1	'左右都截取（都取前面）（保留关键字）
		GetContent=Right(str,Len(str)-Instr(lcase(str),lcase(start))+1)
		GetContent=Left(GetContent,Instr(lcase(GetContent),lcase(last))+Len(last)-1)
		case 2	'只往右截取（取前面的）（去除关键字）
		GetContent=Right(str,Len(str)-Instr(lcase(str),lcase(start))-Len(start)+1)
		case 3	'只往右截取（取前面的）（包含关键字）
		GetContent=Right(str,Len(str)-Instr(lcase(str),lcase(start))+1)
		case 4	'只往左截取（取后面的）（包含关键字）
		GetContent=Left(str,InstrRev(lcase(str),lcase(start))+Len(start)-1)
		case 5	'只往左截取（取后面的）（去除关键字）
		GetContent=Left(str,InstrRev(lcase(str),lcase(start))-1)
		case 6	'只往左截取（取前面的）（包含关键字）
		GetContent=Left(str,Instr(lcase(str),lcase(start))+Len(start)-1)
		case 7	'只往右截取（取后面的）（包含关键字）
		GetContent=Right(str,Len(str)-InstrRev(lcase(str),lcase(start))+1)
		case 8	'只往左截取（取前面的）（去除关键字）
		GetContent=Left(str,Instr(lcase(str),lcase(start))-1)
		case 9	'只往右截取（取后面的）（包含关键字）
		GetContent=Right(str,Len(str)-InstrRev(lcase(str),lcase(start)))
		end select
	Else
		GetContent=""
	End if
	
End function

'过滤空格 回车 制表符
Function filtrate(str)
	str=replace(str,chr(13),"")
	str=replace(str,chr(10),"")
	str=replace(str,chr(9),"")
	filtrate=str
End Function

Function toUTF8(szInput)
	Dim wch, uch, szRet
	Dim x
	Dim nAsc, nAsc2, nAsc3

	If szInput = "" Then
	toUTF8 = szInput
	Exit Function
	End If
	For x = 1 To Len(szInput)
	  wch = Mid(szInput, x, 1)
	  nAsc = AscW(wch)
	  If nAsc < 0 Then nAsc = nAsc + 65536

	  If (nAsc And &HFF80) = 0 Then
		 szRet = szRet & wch
	  Else
		  If (nAsc And &HF000) = 0 Then
			 uch = "%" & Hex(((nAsc \ 2 ^ 6)) or &HC0) & Hex(nAsc And &H3F or &H80)
			 szRet = szRet & uch
		   Else
			  uch = "%" & Hex((nAsc \ 2 ^ 12) or &HE0) & "%" & _
			  Hex((nAsc \ 2 ^ 6) And &H3F or &H80) & "%" & _
			  Hex(nAsc And &H3F or &H80)
			  szRet = szRet & uch
		   End If
	  End If
	Next

	toUTF8 = szRet
End Function

' 使用xmlhttp的方法来获得图片的内容
Function GetRemoteFile(url)
	On Error Resume Next
	Dim httpobjs
	Set httpobjs = Server.CreateObject("Microsoft.XMLHTTP")
	httpobjs.Open "GET", url, False
	httpobjs.Send()
	If httpobjs.readystate<>4 Then 
		Exit Function
	End If
	GetRemoteFile = httpobjs.responseBody
	Set httpobjs = Nothing
	If err.number<>0 Then err.Clear 
End Function

' *********************************
'	保存远程文件到本地硬盘
' *********************************
Function SaveImage(from, tofile)
	dim geturl,objStream,imgs
	geturl=trim(from)
	imgs=GetRemoteFile(geturl)'取得图片的具休内容的过程
	Set objStream = Server.CreateObject("ADODB.Stream")'建立ADODB.Stream对象，必须要ADO 2.5以上版本
	objStream.Type =1'以二进制模式打开
	objStream.Open
	objstream.write imgs '将字符串内容写入缓冲
	objstream.SaveToFile tofile, 2 '-将缓冲的内容写入文件
	objstream.Close() '关闭对象
	Set objstream = Nothing
End Function

Function saveAndReplaceRemoteImg(getContent,savePath,daysPath,imgPreName)
	retstr=initStr(getContent)
	arrimg=split(retstr,"||")'分割字串，取得里面地址列表
	allimg=""
	newimg=""	
	for i=1 to ubound(arrimg)
		if arrimg(i)<>"" and instr(allimg,arrimg(i))<1 then'看这个图片是否已经下载过
			fname=cstr(i&mid(arrimg(i),instrrev(arrimg(i),".")))
			call saveimage(arrimg(i),savePath&imgPreName&"_"&fname)'保存地址的函数，过程见上面
			allimg=allimg&"||"&arrimg(i)'把保存下来的图片的地址串回起来，以确定要替换的地址
			newimg=newimg&"||"&SiteUrl&"/"&Imgsavedir&"/"&dayspath&"/"&imgPreName&"_"&fname'把本地的地址串回起来
		end if
	next
	arrnew=split(newimg,"||")'取得原来的图片地址列表
	arrall=split(allimg,"||")'取得已经保存下来的图片的地址列表
	for i=1 to ubound(arrnew)'执行循环替换原来的地址
		getContent=replace(getContent,arrall(i),arrnew(i))
	next
	saveAndReplaceRemoteImg=getContent
End Function

function getPathList(pathName)
 dim FSO,ServerFolder,getInfo,getInfos,tempS
 getInfo=""
		Set FSO=Server.CreateObject("Scripting.FileSystemObject")
		
		Set ServerFolder=FSO.GetFolder(Server.MapPath(pathName))
			Dim ServerFolderList,ServerFolderEvery
			Set ServerFolderList=ServerFolder.SubFolders
			tempS=""
			For Each ServerFolderEvery IN ServerFolderList
                getInfo=getInfo&tempS&ServerFolderEvery.Name
                tempS="*"
			Next
            getInfo=getInfo&"|"
			Dim ServerFileList,ServerFileEvery
			Set ServerFileList=ServerFolder.Files
			tempS=""
			For Each ServerFileEvery IN ServerFileList
                getInfo=getInfo&tempS&ServerFileEvery.Name
                tempS="*"
			Next
	Set FSO=Nothing
	getInfos=split(getInfo,"|")
	getPathList=getInfos
end function

Function StartURLCache()
		'Declare variables
	Dim strResponse
	Dim objTextStream
	Dim objHTTP
	Dim strDynURL, strDynURL2
	Dim cache_FILEOBJ, cache_FILETXT

	'Get URL
	strDynURL = Request.ServerVariables("URL") & _
	"?" & Request.ServerVariables("QUERY_STRING")
	strDynURL2 = strDynURL

	'Generate file name based on URL
	strDynURL = Replace(strDynURL, "&", "_")
	strDynURL = Replace(strDynURL, "=", "-")
	strDynURL = Replace(strDynURL, "/", "")
	strDynURL = Replace(strDynURL, "?", "$")
	strDynURL = Replace(strDynURL, ".", "@")

	'Create file system object
	set cache_FILEOBJ = Server.CreateObject("Scripting.FileSystemObject")

	'check the "expires" file, which tells us if we've tried to start caching
	strCacheExpiresFilePath = Server.MapPath(".") & "\cachefiles\" & strDynURL & "_started.txt"
	strCacheFilePath = Server.MapPath(".") & "\cachefiles\" & strDynURL & ".txt"
	if not cache_FILEOBJ.FileExists(strCacheExpiresFilePath) then	
		Set cache_FILETXT = cache_FILEOBJ.CreateTextFile(strCacheExpiresFilePath, True)

		'grab the contents of the page
		Set objHTTP = Server.CreateObject("MSXML2.XMLHTTP")
		objHTTP.open "GET", config("siteurl") & strDynURL2, false
		objHTTP.send()
		'output results to screen 
		response.write objHTTP.responseText

		'create cache text file
		Set cache_FILETXT = _
		cache_FILEOBJ.CreateTextFile(strCacheFilePath, True)

		'save results
		cache_FILETXT.WriteLine objHTTP.responseText

		'clean up
		cache_FILETXT.Close()
		set objHTTP = nothing
		response.end
		StartURLCache = TRUE
	Else
		If cache_FILEOBJ.FileExists(strCacheFilePath) then
		  set cache_FILETXT = cache_FILEOBJ.OpenTextFile(strCacheFilePath, 1)
		  response.write cache_FILETXT.ReadAll()
		  cache_FILETXT.Close()
		  response.end
		  StartURLCache = TRUE
		Else
			StartURLCache = FALSE
		End If
	End If
End Function


Function IsValidUrl(url)
        Set xl = Server.CreateObject("Microsoft.XMLHTTP")
        xl.Open "HEAD",url,False
        xl.Send
        IsValidUrl = (xl.status=200)
End Function

