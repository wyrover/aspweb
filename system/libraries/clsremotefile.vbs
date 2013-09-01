Class clsRemoteFile

	
	Private m_strAllowExt
	Private m_nAllowSize
	Private m_strUploadDir
	Private m_strContentPath

	Private sFileExt
	Private sOriginalFileName
	Private sSaveFileName
	Private sPathFileName
	Private nFileNum
	

	Public Property Get AllowExt
		AllowExt = m_strAllowExt
	End Property

	Public Property Let AllowExt(strAllowExt)
		m_strAllowExt = strAllowExt
	End Property

	Public Property Get AllowSize
		AllowSize = m_nAllowSize
	End Property

	Public Property Let AllowSize(nAllowSize)
		m_nAllowSize = nAllowSize
	End Property


	Public Property Get UploadDir
		UploadDir = m_strUploadDir
	End Property

	Public Property Let UploadDir(strUploadDir)
		m_strUploadDir = strUploadDir
	End Property


	Public Property Get ContentPath
		ContentPath = m_strContentPath
	End Property

	Public Property Let ContentPath(strContentPath)
		m_strContentPath = strContentPath
	End Property



	

	' 类初始化
	Private Sub Class_Initialize()

		On Error Resume Next
		m_strAllowExt = "GIF|JPG|PNG|BMP"
		m_nAllowSize = 800
		m_strUploadDir = "attachments/" & Request.ServerVariables("http_host") & "/"
		sPathFileName = "/asp_dc/"
		m_strContentPath = "/attachments/" & Request.ServerVariables("http_host") & "/"
	End Sub

	' 类释放
	Private Sub Class_Terminate()		

	End Sub



	' 取随机文件名
Function GetRndFileName(sExt)
	Dim sRnd
	Randomize
	sRnd = Int(900 * Rnd) + 100
	GetRndFileName = year(now) & month(now) & day(now) & hour(now) & minute(now) & second(now) & sRnd & "." & sExt
End Function



' 转为根路径格式
Function RelativePath2RootPath(url)
	Dim sTempUrl
	sTempUrl = url
	If Left(sTempUrl, 1) = "/" Then
		RelativePath2RootPath = sTempUrl
		Exit Function
	End If

	Dim sWebEditorPath
	sWebEditorPath = Request.ServerVariables("SCRIPT_NAME")
	sWebEditorPath = Left(sWebEditorPath, InstrRev(sWebEditorPath, "/") - 1)
	Do While Left(sTempUrl, 3) = "../"
		sTempUrl = Mid(sTempUrl, 4)
		sWebEditorPath = Left(sWebEditorPath, InstrRev(sWebEditorPath, "/") - 1)
	Loop
	RelativePath2RootPath = sWebEditorPath & "/" & sTempUrl
End Function

' 根路径转为带域名全路径格式
Function RootPath2DomainPath(url)
	Dim sHost, sPort
	sHost = Split(Request.ServerVariables("SERVER_PROTOCOL"), "/")(0) & "://" & Request.ServerVariables("HTTP_HOST")
	sPort = Request.ServerVariables("SERVER_PORT")
	If sPort <> "80" Then
		sHost = sHost & ":" & sPort
	End If
	RootPath2DomainPath = sHost & url
End Function

'================================================
'作  用：替换字符串中的远程文件为本地文件并保存远程文件
'参  数：
'	sHTML		: 要替换的字符串
'	sExt		: 执行替换的扩展名
'================================================
Function ReplaceRemoteUrl(sHTML)
	Dim s_Content
	s_Content = sHTML
	If IsObjInstalled("Microsoft.XMLHTTP") = False then
		ReplaceRemoteUrl = s_Content
		Exit Function
	End If
	
	Dim re, RemoteFile, RemoteFileurl, SaveFileName, SaveFileType
	Set re = new RegExp
	re.IgnoreCase  = True
	re.Global = True
	re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\){1}(([A-Za-z0-9_-])+[.]){1,}(net|com|cn|org|cc|tv|[0-9]{1,3})(\S*\/)((\S)+[.]{1}(" & m_strAllowExt & ")))"

	Set RemoteFile = re.Execute(s_Content)
	Dim a_RemoteUrl(), n, i, bRepeat
	n = 0
	' 转入无重复数据
	For Each RemoteFileurl in RemoteFile
		If n = 0 Then
			n = n + 1
			Redim a_RemoteUrl(n)
			a_RemoteUrl(n) = RemoteFileurl
		Else
			bRepeat = False
			For i = 1 To UBound(a_RemoteUrl)
				If UCase(RemoteFileurl) = UCase(a_RemoteUrl(i)) Then
					bRepeat = True
					Exit For
				End If
			Next
			If bRepeat = False Then
				n = n + 1
				Redim Preserve a_RemoteUrl(n)
				a_RemoteUrl(n) = RemoteFileurl
			End If
		End If		
	Next

	
	' 开始替换操作
	nFileNum = 0
	For i = 1 To n
		SaveFileType = Mid(a_RemoteUrl(i), InstrRev(a_RemoteUrl(i), ".") + 1)
		SaveFileName = GetRndFileName(SaveFileType)
		
		If SaveRemoteFile(SaveFileName, a_RemoteUrl(i)) = True Then
			
			nFileNum = nFileNum + 1
			If nFileNum > 0 Then
				sOriginalFileName = sOriginalFileName & "|"
				sSaveFileName = sSaveFileName & "|"
				sPathFileName = sPathFileName & "|"
			End If
			sOriginalFileName = sOriginalFileName & Mid(a_RemoteUrl(i), InstrRev(a_RemoteUrl(i), "/") + 1)
			sSaveFileName = sSaveFileName & SaveFileName
			sPathFileName = sPathFileName & m_strContentPath & SaveFileName
			s_Content = Replace(s_Content, a_RemoteUrl(i), m_strContentPath & SaveFileName, 1, -1, 1)
		End If
	Next

	ReplaceRemoteUrl = s_Content
End Function

'================================================
'作  用：保存远程的文件到本地
'参  数：s_LocalFileName ------ 本地文件名
'		 s_RemoteFileUrl ------ 远程文件URL
'返回值：True  ----成功
'        False ----失败
'================================================
Function SaveRemoteFile(s_LocalFileName, s_RemoteFileUrl)
	Dim Ads, Retrieval, GetRemoteData
	Dim bError
	bError = False
	SaveRemoteFile = False
	'On Error Resume Next
	Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
	With Retrieval
		.Open "Get", s_RemoteFileUrl, False, "", ""
		.Send
		GetRemoteData = .ResponseBody
	End With
	Set Retrieval = Nothing

	If LenB(GetRemoteData) > m_nAllowSize*1024 Then		
		bError = True
		Response.Write Server.MapPath(m_strUploadDir & s_LocalFileName)
	Else
		Set Ads = Server.CreateObject("Adodb.Stream")
		With Ads
			.Type = 1
			.Open
			.Write GetRemoteData
			.SaveToFile Server.MapPath(m_strUploadDir & s_LocalFileName), 2
			.Cancel()
			.Close()
			
		End With
		Set Ads=nothing
	End If

	If Err.Number = 0 And bError = False Then
		SaveRemoteFile = True
	Else
		Err.Clear
	End If
End Function

'================================================
'作  用：检查组件是否已经安装
'参  数：strClassString ----组件名
'返回值：True  ----已经安装
'        False ----没有安装
'================================================
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

End Class


