Function isPostBack()
    isPostBack = False
    If Request.ServerVariables("REQUEST_METHOD") = "POST" Then isPostBack = True
End Function

'判断提交信息是否来自外部
Public Function ChkIsOuter()
	Dim server_v1,server_v2
	ChkIsOuter=True 
	server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
	server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
	If Mid(server_v1,8,len(server_v2))=server_v2 Then ChkIsOuter=False 
End Function

' ********************************************
' *	 判断组件是否安装
' *	 @classString		字符串
' *	 @return				布尔值
' ********************************************
Function IsObjectInstalled(classString)
	On Error Resume Next

	IsObjectInstalled = False
	Err = 0
	Dim objTest

	objTest = Server.CreateObject(classString)
	If Err = 0 Or Err = -2147352567 Then
		IsObjectInstalled = True
	End If
	
	Set objTest = Nothing
	Err = 0
	
End Function


Function IsObj(x)
	If Not IsDll(x) Then
		IsObj="<font color='red'><b>×</b></font>"
	Else
		IsObj="<b>√</b>&nbsp;" & getver(x)
	End If
End Function

Function IsDll(strClassString)
	On Error Resume Next
	IsDll=False
	Err = 0
	Dim xTestObj
	Set xTestObj=CreateObject(strClassString)
	If 0 = Err Then IsDll=True
	Set xTestObj=Nothing
	Err = 0
End Function

Function getver(Classstr)
On Error Resume Next
ver=""
Err = 0
Dim Obj
Set Obj=CreateObject(Classstr)
If 0 = Err Then ver=obj.version
Set Obj=Nothing
Err = 0
End Function

' ********************************************
' *	IIf条件表达式
' ********************************************
Function IIf(condition, resTrue, resFalse)
	If condition Then
		IIf = resTrue
	Else
		IIf = resFalse
	End if
End Function

' ******************************************
' * 获取用户IP
' * @return				字符串
' ******************************************
Function GetClientIP()
    Dim ip

    '如果客户端用了代理服务器，则应该用ServerVariables("HTTP_X_FORWARDED_FOR")方法
    ip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
    If ip = "" Or IsNull(GetClientIP) Or IsEmpty(GetClientIP) Then
    '如果客户端没用代理，应该用Request.ServerVariables("REMOTE_ADDR")方法
    ip = Request.ServerVariables("REMOTE_ADDR")
    End If

    GetClientIP = ip
End Function



' *****************************************
' * 获取当前路径
' * @return				字符串
' *****************************************
Function GetCurrentPath()
	Dim TempPath,Path
	TempPath = Request.ServerVariables("Path_info")
	Path = Left(TempPath,InstrRev(TempPath,"/"))
	GetCurrentPath = GetDoMain & Path
End Function

' ****************************************
' *	 据时间返回Image Html代码
' * @mytime			时间
' ****************************************
Function GetImage(ByVal mytime)

	Dim nowtime

	nowtime = Now()
	if (datediff("d", mytime, nowtime)) < 1 Then
		GetImage = "<img src=""images/new.gif"">"
	End If

End Function

' ***********************************
'	获取文件图标
' ***********************************
Function GetFileIcons(str)

    Dim FileIcon, Target

    Select Case str
        Case ".jpg"
        FileIcon = "jpg.gif"
        Case ".gif"
        FileIcon = "gif.gif"
        Case ".bmp"
        FileIcon = "bmp.gif"
        Case ".png"
        FileIcon = "png.gif"
        Case ".zip"
        FileIcon = "zip.gif"
        Case ".rar"
        FileIcon = "rar.gif"
        Case ".swf"
        FileIcon = "swf.gif"
        Case ".mdb"
        FileIcon = "mdb.gif"
        Case ".doc"
        FileIcon = "doc.gif"
        Case ".xls"
        FileIcon = "xls.gif"
        Case ".pdf"
        FileIcon = "pdf.gif"
        Case ".mbk"
        FileIcon = "mbk.gif"
        Case ".mp3"
        FileIcon = "mp3.gif"
        Case ".wmv"
        FileIcon = "wma.gif"
        Case ".wma"
        FileIcon = "wma.gif"
        Case Else
        FileIcon = "unknow.gif"
    End Select

    GetFileIcons = "<img src=""images/file/" & FileIcon & """/>"
 
End Function



' *******************************************
' * 防刷新
' *******************************************
Sub Frash()
    Dim AppealNum, AppealCount
    AppealNum = 10 '同一IP60秒内请求限制10次
    AppealCount = Request.cookies("AppealCount")
    If AppealCount = "" Then
    response.cookies("AppealCount") = 1
    AppealCount = 1
    response.cookies("AppealCount").expires = DateAdd("s", 60, Now())
    Else
    response.cookies("AppealCount") = AppealCount + 1
    response.cookies("AppealCount").expires = DateAdd("s", 60, Now())
    End If
    If Int(AppealCount) > Int(AppealNum) Then
    response.write "抓取很累，歇一会儿吧！"
    response.End
    End If
End Sub


Function GetPathInfo()
	GetPathInfo = Request.ServerVariables("SCRIPT_NAME") 
    'Request.ServerVariables("PATH_INFO") 
    'Request.ServerVariables("URL")
End Function

Function GetPageFileName()
	scr = Request.ServerVariables("SCRIPT_NAME") & "<br>" 
    loc = instrRev(scr,"/") 
    scr = mid(scr, loc+1, len(scr) - loc) 
    GetPageFileName = scr
End Function

Function GetRunDir()
	Dim scriptname'获得目录
	scriptname=request.servervariables("script_name")
	page=replace(scriptname,"\","/")
	page=lcase(right(page,len(page)-instrrev(page,"/")))
	systempath=left(scriptname,len(scriptname)-len(page)-1)
	systempath=right(systempath,len(systempath)-instrrev(systempath,"/"))
	GetRunDir = systempath
End Function




Public Function Checkin(s) 
		s = trim(s) 
		s = replace(s," ","&amp;nbsp;") 
		s = replace(s,"'","&amp;#39;") 
		s = replace(s,"""","&amp;quot;") 
		s = replace(s,"&lt;","&amp;lt;") 
		s = replace(s,"&gt;","&amp;gt;") 
		Checkin=s 
End Function 

Public Sub DelFiles(delfilesname,filespath) 
		Dim FileDelete,files,strFileFullPath,filesNum		
		If Right(filespath,1)<>"\" Then filespath = filespath & "\"
		If delfilesname<>"" And Not IsNull(delfilesname) Then
			Set FileDelete = CreateObject("Scripting.FileSystemObject")
			files = Split(delfilesname & "|","|")
			For filesNum=0 to Ubound(files)-1
				strFileFullPath = filespath + files(filesNum)
				If FileDelete.FileExists(strFileFullPath) Then FileDelete.DeleteFile(strFileFullPath)
			Next
		End If
End Sub

'================================================
' 函数名：RootPath2DomainPath
' 作  用：根路径转为带域名全路径格式
' 参  数：url ----原URL
' 返回值：转换后的URL
'================================================
Function RootPath2DomainPath(url)
	Dim sHost, sPort
	sHost = Split(LCase(Request.ServerVariables("SERVER_PROTOCOL")), "/")(0) & "://" & Request.ServerVariables("HTTP_HOST")
	sPort = Request.ServerVariables("SERVER_PORT")
	If sPort <> "80" Then
		sHost = sHost & ":" & sPort
	End If
	RootPath2DomainPath = sHost & url
End Function

'================================================
' 函数名：RelativePath2RootPath
' 作  用：转为根路径格式
' 参  数：url ----原URL
' 返回值：转换后的URL
'================================================
Function RelativePath2RootPath(url)
	Dim sTempUrl
	sTempUrl = url
	If Left(sTempUrl, 1) = "/" Then
		RelativePath2RootPath = sTempUrl
		Exit Function
	End If

	Dim sFilePath
	sFilePath = Request.ServerVariables("SCRIPT_NAME")
	sFilePath = Left(sFilePath, InstrRev(sFilePath, "/") - 1)
	Do While Left(sTempUrl, 3) = "../"
		sTempUrl = Mid(sTempUrl, 4)
		sFilePath = Left(sFilePath, InstrRev(sFilePath, "/") - 1)
	Loop
	RelativePath2RootPath = sFilePath & "/" & sTempUrl
End Function

'================================================
' 函数名：CreatePath
' 作  用：按月份自动创建文件夹
' 参  数：fromPath ----原文件夹路径
'================================================
Function CreatePath(fromPath)
	Dim objFSO, uploadpath
	uploadpath = Year(Now) & "-" & Month(Now) '以年月创建上传文件夹，格式：2003－8
	On Error Resume Next
	Set objFSO = CreateObject(Scripting.FileSystemObject)
	If objFSO.FolderExists(Server.MapPath(fromPath & uploadpath)) = False Then
		objFSO.CreateFolder Server.MapPath(fromPath & uploadpath)
	End If
	If Err.Number = 0 Then
		CreatePath = uploadpath & "/"
	Else
		CreatePath = ""
	End If
	Set objFSO = Nothing
End Function

' Convert a string to a date or datetime
' IN  : sDate (string) : source (format YYYYMMDD HH:MM:SS or YYYYMMDD)
' OUT : (datetime) : destination
Function StringToDate(strDate)
	Dim dDate, sDate

	sDate = trim(strDate)
	select case Len(sDate)
		case 17
			dDate = DateSerial(Left(sDate, 4), Mid(sDate, 5, 2), Mid(sDate, 7, 2)) + TimeSerial(Mid(sDate, 10, 2), Mid(sDate, 13, 2), Mid(sDate, 16, 2))
		case 8
			dDate = DateSerial(Left(sDate, 4), Mid(sDate, 5, 2), Mid(sDate, 7, 2))
		case else
			if isDate(sDate) Then
				dDate = CDate(sDate)
			end if
	End select
	StringToDate = dDate
End Function


' ********************************************
' *	 转换XML
' * @sourceFile		源文件
' *	 @styleFile		样式表文件
' *	 @return			字符串
' ********************************************
Function getXML(sourceFile, styleFile)
	Dim source, style, xmlhttp

	Set xmlhttp = Server.CreateObject("Microsoft.XMLHTTP")

	xmlhttp.Open "GET", sourceFile, false
	xmlhttp.send()
	
	set source = Server.CreateObject("Microsoft.XMLDOM")
	source.async = false
	source.loadxml(xmlhttp.responseXML.xml)

	xmlhttp.Open "GET", styleFile, false
	xmlhttp.send()

	set style = Server.CreateObject("Microsoft.XMLDOM")
	style.async = false
	style.loadxml(xmlhttp.responseXML.xml)

	getXML = replace(source.transformNode(style),"&apos;","&#39;")

	set source = nothing
	set style = nothing
	set xmlhttp = nothing
End Function

' ********************************************
' *	 取随机数
' * @lowerbound		下限
' *	 @upperbound		上限
' *	 @return				字符串
' ********************************************
Function CreateRandomNumber(lowerbound, upperbound)
	Randomize
	CreateRandomNumber = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
End Function

Function CreatePassword(numchar)
	Dim avail, parola, f, i

	avail = "abcdefghijklmnopqrstuvwxyz1234567890"
	Randomize
	parola = ""
	for f = 1 to numchar
		i = (CInt(len(avail) * Rnd + 1) mod len(avail)) + 1
		parola = parola & mid(avail, i, 1)
	next
	CreatePassword = parola
End Function

' **************************************************
' * 转换为有效的 ID
' * @strS					字符串
' * @return				Integer (>=0)
' **************************************************
Function CID(strS)
	Dim intI
	intI = 0
	If IsNull(strS) Or strS = "" Then
		intI = 0
	Else
		If Not IsNumeric(strS) Then
			intI = 0
		Else
			Dim intk
			On Error Resume Next
			intk = Abs(Clng(strS))
			If Err.Number = 6 Then intk = 0  ''数据溢出
			Err.Clear
			intI = intk
		End If
	End If
	CID = intI
End Function

' **************************************************
' * 判断用户名是否合法
' * @uName					字符串
' * @return				True Or False
' **************************************************
Function IstrueName(uName)
	Dim Hname,i
	IstrueName = False
	Hname = Array("=","%",chr(32),"?","&",";",",","'",",",chr(34),chr(9),"","$","|")
	For i = 0 To Ubound(Hname)
		If InStr(uName,Hname(i)) > 0 Then
			Exit Function
		End If
	Next
	IstrueName=True 
End Function

' **************************************************
' * 正则表达式替换
' * @fString					字符串
' * @patrn					模式
' * @replStr					替换字符串
' * @return				
' **************************************************
Public Function ReplaceText(fString,patrn,replStr)
	Dim regEx
	Set regEx = New RegExp  ' 建立正则表达式。
	regEx.Pattern = patrn   ' 设置模式。
	regEx.IgnoreCase = True ' 设置是否区分大小写。
	regEx.Global = True     ' 设置全局可用性。 
	ReplaceText = regEx.Replace(fString, replStr) ' 作替换。
	regEx=Null:Set regEx=Nothing
End Function


'**************************************************
'函数名：HTMLEncode
'作  用：过虑字符
'参  数：str-----要过虑的字符
'返回值：过虑后的字符
'**************************************************
Public Function HTMLEncode(fString)
	If fString="" or IsNull(fString) Then 
		Exit Function
	Else
		If Instr(fString,"'")>0 Then 
			fString = replace(fString, "'","&#39;")
		End If
		fString = replace(fString, ">", "&gt;")
		fString = replace(fString, "<", "&lt;")
		fString = Replace(fString, CHR(32), "&nbsp;")
		fString = Replace(fString, CHR(9), "&nbsp;")
		fString = Replace(fString, CHR(34), "&quot;")
		fString = Replace(fString, CHR(13),"")
		fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
		fString = Replace(fString, CHR(10), "<BR>")
		fString = Replace(fString, CHR(39), "&#39;")
		fString = Replace(fString, CHR(0), "")
		fString = ChkBadWords(fString)
		HTMLEncode = fString
	End If
End Function

'还原字符处理
Public Function iHTMLEncode(fString)
	If fString="" or IsNull(fString) Then 
		Exit Function
	Else
		If Instr(fString,"'")>0 Then 
			fString = replace(fString, "'","&#39;")
		End If
		fString = replace(fString, "&gt;"	, ">")
		fString = replace(fString, "&lt;"	, "<")
		fString = Replace(fString, "&nbsp;"	, CHR(32))
		fString = Replace(fString, "&nbsp;"	, CHR(9))
		fString = Replace(fString, "&quot;"	, CHR(34))
		fString = Replace(fString, ""		, CHR(13))
		fString = Replace(fString, "</P><P>", CHR(10) & CHR(10))
		fString = Replace(fString, "<BR>"	, CHR(10))
		fString = Replace(fString, ""		, CHR(0))
		fString = Replace(fString, "&#39;"	, CHR(39))
		fString = ChkBadWords(fString)
		iHTMLEncode = fString
	End If
End Function

'**************************************************
'函数名：Emailto
'作  用：发送邮件
'参  数：--
'返回值：--
'**************************************************
Sub Emailto(a,b,c)
	if Not IsValidEmail(a) Then team.error2 " E-mail地址填写错误！"
	Select Case team.Forum_setting(1)
		Case "1"
			Dim JMail
			Set JMail=Server.CreateObject("JMail.Message")
			If -2147221005 = Err Then team.error2 "本服务器不支持 JMail.Message 组件！"
			JMail.Charset="gb2312"
			JMail.AddRecipient a
			JMail.Subject = b
			JMail.Body = c
			JMail.From = team.Forum_setting(57)						'发送人地址
			JMail.MailServerUserName = team.Forum_setting(41)		'服务器登陆用户名
			JMail.MailServerPassword = team.Forum_setting(55)		'服务器登陆密码
			JMail.Send team.Forum_setting(58)						'服务器地址
			Set JMail=nothing
		Case "2"
			Dim MailObject
			Set MailObject = Server.CreateObject("CDONTS.NewMail")
			If -2147221005 = Err Then team.error2 "本服务器不支持 CDONTS.NewMail 组件！"
			MailObject.Send team.Forum_setting(57),a,b,c
			'MailObject.Send "发送方邮件地址","接收方邮件地址","主题","邮件正文"
			Set MailObject=nothing
		Case Else
			Exit Sub
	End Select
End Sub

Function IsValidEmail(email)
	Dim names, name, i, c
	IsValidEmail = True
	names = Split(email, "@")
	If UBound(names) <> 1 Then
		IsValidEmail = False
		Exit Function
	End If
	For Each name In names
		If Len(name) <= 0 Then
			IsValidEmail = False
			Exit Function
		End If
		For i = 1 To Len(name)
			c = Lcase(Mid(name, i, 1))
			If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) Then
				IsValidEmail = False
				Exit Function
			End If
		Next
		If Left(name, 1) = "." or Right(name, 1) = "." Then
			IsValidEmail = False
			Exit Function
		End If
	Next
	If InStr(names(1), ".") <= 0 Then
		IsValidEmail = False
		Exit Function
	End If
	i = Len(names(1)) - InStrRev(names(1), ".")
	If i <> 2 and i <> 3 Then
		IsValidEmail = False
		Exit Function
	End If
	If InStr(email, "..") > 0 Then
		IsValidEmail = False
	End If
End Function

'*************************************
'日期转换函数
'*************************************
Function DateToStr(DateTime,ShowType)  
	Dim DateMonth,DateDay,DateHour,DateMinute,DateWeek,DateSecond
	Dim FullWeekday,shortWeekday,Fullmonth,Shortmonth,TimeZone1,TimeZone2
	TimeZone1="+0800"
	TimeZone2="+08:00"
	FullWeekday=Array("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday")
	shortWeekday=Array("Sun","Mon","Tue","Wed","Thu","Fri","Sat")
    Fullmonth=Array("January","February","March","April","May","June","July","August","September","October","November","December")
    Shortmonth=Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")

	DateMonth=Month(DateTime)
	DateDay=Day(DateTime)
	DateHour=Hour(DateTime)
	DateMinute=Minute(DateTime)
	DateWeek=weekday(DateTime)
	DateSecond=Second(DateTime)
	If Len(DateMonth)<2 Then DateMonth="0"&DateMonth
	If Len(DateDay)<2 Then DateDay="0"&DateDay
	If Len(DateMinute)<2 Then DateMinute="0"&DateMinute
	Select Case ShowType
	Case "Y-m-d"  
		DateToStr=Year(DateTime)&"-"&DateMonth&"-"&DateDay
	Case "Y-m-d H:I A"
		Dim DateAMPM
		If DateHour>12 Then 
			DateHour=DateHour-12
			DateAMPM="PM"
		Else
			DateHour=DateHour
			DateAMPM="AM"
		End If
		If Len(DateHour)<2 Then DateHour="0"&DateHour	
		DateToStr=Year(DateTime)&"-"&DateMonth&"-"&DateDay&" "&DateHour&":"&DateMinute&" "&DateAMPM
	Case "Y-m-d H:I:S"
		If Len(DateHour)<2 Then DateHour="0"&DateHour	
		If Len(DateSecond)<2 Then DateSecond="0"&DateSecond
		DateToStr=Year(DateTime)&"-"&DateMonth&"-"&DateDay&" "&DateHour&":"&DateMinute&":"&DateSecond
	Case "YmdHIS"
		DateSecond=Second(DateTime)
		If Len(DateHour)<2 Then DateHour="0"&DateHour	
		If Len(DateSecond)<2 Then DateSecond="0"&DateSecond
		DateToStr=Year(DateTime)&DateMonth&DateDay&DateHour&DateMinute&DateSecond	
	Case "ym"
		DateToStr=Right(Year(DateTime),2)&DateMonth
	Case "d"
		DateToStr=DateDay
    Case "ymd"
        DateToStr=Right(Year(DateTime),4)&DateMonth&DateDay
    Case "mdy" 
        Dim DayEnd
        select Case DateDay
         Case 1 
          DayEnd="st"
         Case 2
          DayEnd="nd"
         Case 3
          DayEnd="rd"
         Case Else
          DayEnd="th"
        End Select 
        DateToStr=Fullmonth(DateMonth-1)&" "&DateDay&DayEnd&" "&Right(Year(DateTime),4)
    Case "w,d m y H:I:S" 
		DateSecond=Second(DateTime)
		If Len(DateHour)<2 Then DateHour="0"&DateHour	
		If Len(DateSecond)<2 Then DateSecond="0"&DateSecond
        DateToStr=shortWeekday(DateWeek-1)&","&DateDay&" "& Left(Fullmonth(DateMonth-1),3) &" "&Right(Year(DateTime),4)&" "&DateHour&":"&DateMinute&":"&DateSecond&" "&TimeZone1
    Case "y-m-dTH:I:S"
		If Len(DateHour)<2 Then DateHour="0"&DateHour	
		If Len(DateSecond)<2 Then DateSecond="0"&DateSecond
		DateToStr=Year(DateTime)&"-"&DateMonth&"-"&DateDay&"T"&DateHour&":"&DateMinute&":"&DateSecond&TimeZone2
	Case Else
		If Len(DateHour)<2 Then DateHour="0"&DateHour
		DateToStr=Year(DateTime)&"-"&DateMonth&"-"&DateDay&" "&DateHour&":"&DateMinute
	End Select
End Function

'*************************************
'计算随机数
'*************************************
function randomStr(intLength)
    dim strSeed,seedLength,pos,str,i
    strSeed = "abcdefghijklmnopqrstuvwxyz1234567890"
    seedLength=len(strSeed)
    str=""
    Randomize
    for i=1 to intLength
     str=str+mid(strSeed,int(seedLength*rnd)+1,1)
    next
    randomStr=str
end function

'*************************************
'过滤文件名字
'*************************************
Function FixName(UpFileExt)
	If IsEmpty(UpFileExt) Then Exit Function
	FixName = Ucase(UpFileExt)
	FixName = Replace(FixName,Chr(0),"")
	FixName = Replace(FixName,".","")
	FixName = Replace(FixName,"ASP","")
	FixName = Replace(FixName,"ASA","")
	FixName = Replace(FixName,"ASPX","")
	FixName = Replace(FixName,"CER","")
	FixName = Replace(FixName,"CDX","")
	FixName = Replace(FixName,"HTR","")
End Function

'*************************************
'限制上传文件类型
'*************************************  
Function IsvalidFile(File_Type)
	IsvalidFile = False
	Dim GName
	Dim arrFileType
	arrFileType = Split(config("upload_file_ext"), "|")
	For Each GName in arrFileType		
		If LCase(File_Type) = LCase(GName) Then
			IsvalidFile = True
			Exit For
		End If
	Next
End Function

'*************************************
' 跟踪输出
'*************************************  
Function Trace(s)
    On Error Resume Next
    If IsArray(s) Then
        For i = 0 To UBound(s)
            Response.Write(s(i) & vbCrLf)
        Next
    Else
        Response.Write(s & vbCrLf)
    End If
    Response.End()
End Function

' 用表格显示记录集getrows生成的数组的表结构 
Function ShowRsArr(rsArr)
showHtml="<table width=100% border=1 cellspacing=0 cellpadding=0>" 
    If Not IsEmpty(rsArr) Then 
        For y=0 To Ubound(rsArr,2) 
        showHtml=showHtml&"<tr>" 
            for x=0 to Ubound(rsArr,1) 
                showHtml=showHtml& "<td>"&rsArr(x,y)&"</td>" 
            next 
        showHtml=showHtml&"</tr>" 
        next 
    Else 
        RshowHtml=showHtml&"<tr>" 
        showHtml=showHtml&"<td>No Records</td>" 
        showHtml=showHtml&"</tr>" 
    End If 
        showHtml=showHtml&"</table>" 
    ShowRsArr=showHtml
End Function 


'取字段数据每个汉字的拼音首字母
Function getpychar(char)
    tmp = 65536 + Asc(char)
    If(tmp>= 45217 And tmp<= 45252) Then
        getpychar = "A"
    ElseIf(tmp>= 45253 And tmp<= 45760) Then
        getpychar = "B"
    ElseIf(tmp>= 47761 And tmp<= 46317) Then
        getpychar = "C"
    ElseIf(tmp>= 46318 And tmp<= 46825) Then
        getpychar = "D"
    ElseIf(tmp>= 46826 And tmp<= 47009) Then
        getpychar = "E"
    ElseIf(tmp>= 47010 And tmp<= 47296) Then
        getpychar = "F"
    ElseIf(tmp>= 47297 And tmp<= 47613) Then
        getpychar = "G"
    ElseIf(tmp>= 47614 And tmp<= 48118) Then
        getpychar = "H"
    ElseIf(tmp>= 48119 And tmp<= 49061) Then
        getpychar = "J"
    ElseIf(tmp>= 49062 And tmp<= 49323) Then
        getpychar = "K"
    ElseIf(tmp>= 49324 And tmp<= 49895) Then
        getpychar = "L"
    ElseIf(tmp>= 49896 And tmp<= 50370) Then
        getpychar = "M"
    ElseIf(tmp>= 50371 And tmp<= 50613) Then
        getpychar = "N"
    ElseIf(tmp>= 50614 And tmp<= 50621) Then
        getpychar = "O"
    ElseIf(tmp>= 50622 And tmp<= 50905) Then
        getpychar = "P"
    ElseIf(tmp>= 50906 And tmp<= 51386) Then
        getpychar = "Q"
    ElseIf(tmp>= 51387 And tmp<= 51445) Then
        getpychar = "R"
    ElseIf(tmp>= 51446 And tmp<= 52217) Then
        getpychar = "S"
    ElseIf(tmp>= 52218 And tmp<= 52697) Then
        getpychar = "T"
    ElseIf(tmp>= 52698 And tmp<= 52979) Then
        getpychar = "W"
    ElseIf(tmp>= 52980 And tmp<= 53640) Then
        getpychar = "X"
    ElseIf(tmp>= 53689 And tmp<= 54480) Then
        getpychar = "Y"
    ElseIf(tmp>= 54481 And tmp<= 62289) Then
        getpychar = "Z"
    Else '如果不是中文，则不处理
        getpychar = char
End If
End Function

'*************************************
' 获取拼音
'*************************************  
Function getpy(Str)
    For i = 1 To Len(Str)
        getpy = getpy&getpychar(Mid(Str, i, 1))
    Next
End Function

Function bytes2BSTR(vIn)
    Dim strReturn
    Dim i, ThisCharCode, NextCharCode
    strReturn = ""
    For i = 1 To LenB(vIn)
        ThisCharCode = AscB(MidB(vIn, i, 1))
        If ThisCharCode < &H80 Then
            strReturn = strReturn & Chr(ThisCharCode)
        Else
            NextCharCode = AscB(MidB(vIn, i + 1, 1))
            strReturn = strReturn & Chr(CLng(ThisCharCode) * &H100 + CInt(NextCharCode))
            i = i + 1
        End If
    Next
    bytes2BSTR = strReturn
End Function

'*************************************
' 调试用
'*************************************  
Function Trace(s)
    On Error Resume Next
    If IsArray(s) Then
        For i = 0 To UBound(s)
            Response.Write(s(i) & vbCrLf)
        Next
    Else
        Response.Write(s & vbCrLf)
    End If
    Response.End()
End Function


Public Function GetSearchKeyword(RefererUrl) '搜索关键词
 if RefererUrl="" or len(RefererUrl)<1 then exit function
    
  on error resume next
  
  Dim re
  Set re = New RegExp
  re.IgnoreCase = True
  re.Global = True
  Dim a,b,j
  '模糊查找关键词，此方法速度较快，范围也较大
  re.Pattern = "(word=([^&]*)|q=([^&]*)|p=([^&]*)|query=([^&]*)|name=([^&]*)|_searchkey=([^&]*)|baidu.*?w=([^&]*))"
  Set a = re.Execute(RefererUrl)
  If a.Count>0 then
   Set b = a(a.Count-1).SubMatches
   For j=1 to b.Count
    If Len(b(j))>0 then 
     if instr(1,RefererUrl,"google",1) then 
       GetSearchKeyword=Trim(U8Decode(b(j)))
      elseif instr(1,refererurl,"yahoo",1) then 
       GetSearchKeyword=Trim(U8Decode(b(j)))
      elseif instr(1,refererurl,"yisou",1) then
       GetSearchKeyword=Trim(getkey(b(j)))
      elseif instr(1,refererurl,"3721",1) then
       GetSearchKeyword=Trim(getkey(b(j)))
      else 
       GetSearchKeyword=Trim(getkey(b(j)))
     end if
     Exit Function
    end if
   Next
  End If
  if err then
  err.clear
  GetSearchKeyword = RefererUrl
  else
  GetSearchKeyword = ""  
  end if  
 End Function


 Function URLEncoding(vstrIn)
  dim strReturn,i,thischr
    strReturn = ""
    For i = 1 To Len(vstrIn)
        ThisChr = Mid(vStrIn,i,1)
        If Abs(Asc(ThisChr)) < &HFF Then
            strReturn = strReturn & ThisChr
        Else
            innerCode = Asc(ThisChr)
            If innerCode < 0 Then
                innerCode = innerCode + &H10000
            End If
            Hight8 = (innerCode  And &HFF00)\ &HFF
            Low8 = innerCode And &HFF
            strReturn = strReturn & "%" & Hex(Hight8) &  "%" & Hex(Low8)
        End If
    Next
    URLEncoding = strReturn
End Function
function getkey(key)
dim oreq
set oreq = CreateObject("MSXML2.XMLHTTP")
oReq.open "POST","http://"&WebUrl&"/system/ShowGB2312XML.asp?a="&key,false
oReq.send
getkey=UTF2GB(oReq.responseText)
end function
function chinese2unicode(Str) 
  dim i 
  dim Str_one 
  dim Str_unicode 
  for i=1 to len(Str) 
    Str_one=Mid(Str,i,1) 
    Str_unicode=Str_unicode&chr(38) 
    Str_unicode=Str_unicode&chr(35) 
    Str_unicode=Str_unicode&chr(120) 
    Str_unicode=Str_unicode& Hex(ascw(Str_one)) 
    Str_unicode=Str_unicode&chr(59) 
  next 
  Response.Write Str_unicode 
end function     
  
function UTF2GB(UTFStr)
Dim dig,GBSTR
    for Dig=1 to len(UTFStr)
        if mid(UTFStr,Dig,1)="%" then
            if len(UTFStr) >= Dig+8 then
                GBStr=GBStr & ConvChinese(mid(UTFStr,Dig,9))
                Dig=Dig+8
            else
                GBStr=GBStr & mid(UTFStr,Dig,1)
            end if
        else
            GBStr=GBStr & mid(UTFStr,Dig,1)
        end if
    next
    UTF2GB=GBStr
end function 


function ConvChinese(x) 
dim a,i,j,DigS,Unicode
    A=split(mid(x,2),"%")
    i=0
    j=0
    
    for i=0 to ubound(A) 
        A(i)=c16to2(A(i))
    next
        
    for i=0 to ubound(A)-1
        DigS=instr(A(i),"0")
        Unicode=""
        for j=1 to DigS-1
            if j=1 then 
                A(i)=right(A(i),len(A(i))-DigS)
                Unicode=Unicode & A(i)
            else
                i=i+1
                A(i)=right(A(i),len(A(i))-2)
                Unicode=Unicode & A(i) 
            end if 
        next
        
        if len(c2to16(Unicode))=4 then
            ConvChinese=ConvChinese & chrw(int("&H" & c2to16(Unicode)))
        else
            ConvChinese=ConvChinese & chr(int("&H" & c2to16(Unicode)))
        end if
    next
end function

function U8Decode(enStr)
  '输入一堆有%分隔的字符串，先分成数组，根据utf8规则来判断补齐规则
  '输入:关 E5 85 B3  键  E9 94 AE 字   E5 AD 97
  '输出:关 B9D8  键  BCFC 字   D7D6
  dim c,i,i2,v,deStr,WeiS

  for i=1 to len(enStr)
    c=Mid(enStr,i,1)
    if c="%" then
      v=c16to2(Mid(enStr,i+1,2))
      '判断第一次出现0的位置，
      '可能是1(单字节)，3(3-1字节)，4，5，6，7不可能是2和大于7
      '理论上到7，实际不会超过3。
      WeiS=instr(v,"0")
      v=right(v,len(v)-WeiS)'第一个去掉最左边的WeiS个
      i=i+3
      for i2=2 to WeiS-1
        c=c16to2(Mid(enStr,i+1,2))
        c=right(c,len(c)-2)'其余去掉最左边的两个
        v=v & c
        i=i+3
      next
      if len(c2to16(v)) =4 then
        deStr=deStr & chrw(c2to10(v))
      else
        deStr=deStr & chr(c2to10(v))
      end if
      i=i-1
    else
      if c="+" then
        deStr=deStr&" "
      else
        deStr=deStr&c
      end if
    end if
  next
  U8Decode = deStr
end function

function c16to2(x)
 '这个函数是用来转换16进制到2进制的，可以是任何长度的，一般转换UTF-8的时候是两个长度，比如A9
 '比如：输入“C2”，转化成“11000010”,其中1100是"c"是10进制的12（1100），那么2（10）不足4位要补齐成（0010）。
 dim tempstr
 dim i:i=0'临时的指针

 for i=1 to len(trim(x))
  tempstr= c10to2(cint(int("&h" & mid(x,i,1))))
  do while len(tempstr)<4
   tempstr="0" & tempstr'如果不足4位那么补齐4位数
  loop
  c16to2=c16to2 & tempstr
 next
end function

function c2to16(x)
  '2进制到16进制的转换，每4个0或1转换成一个16进制字母，输入长度当然不可能不是4的倍数了

  dim i:i=1'临时的指针
  for i=1 to len(x)  step 4
   c2to16=c2to16 & hex(c2to10(mid(x,i,4)))
  next
end function

function c2to10(x)
  '单纯的2进制到10进制的转换，不考虑转16进制所需要的4位前零补齐。
  '因为这个函数很有用！以后也会用到，做过通讯和硬件的人应该知道。
  '这里用字符串代表二进制
   c2to10=0
   if x="0" then exit function'如果是0的话直接得0就完事
   dim i:i=0'临时的指针
   for i= 0 to len(x) -1'否则利用8421码计算，这个从我最开始学计算机的时候就会，好怀念当初教我们的谢道建老先生啊！
    if mid(x,len(x)-i,1)="1" then c2to10=c2to10+2^(i)
   next
end function

function c10to2(x)
'10进制到2进制的转换
  dim sign, result
  result = ""
  '符号
  sign = sgn(x)
  x = abs(x)
  if x = 0 then
    c10to2 = 0
    exit function
  end if
  do until x = "0"
    result = result & (x mod 2)
    x = x \ 2
  loop
  result = strReverse(result)
  if sign = -1 then
    c10to2 = "-" & result
  else
    c10to2 = result
  end if
end function


Public Function ArrayToxml(DataArray, Recordset, row, xmlroot)
    Dim i, node, rs, j
    If xmlroot = "" Then xmlroot = "xml"
    Set ArrayToxml = Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
    ArrayToxml.appendChild(ArrayToxml.createElement(xmlroot))
    If row = "" Then row = "row"
    For i = 0 To UBound(DataArray, 2)
        Set Node = ArrayToxml.createNode(1, row, "")
        j = 0
        For Each rs in Recordset.Fields
            node.Attributes.setNamedItem(ArrayToxml.createNode(2, LCase(rs.Name), "")).text = DataArray(j, i)& ""
            j = j + 1
        Next
        ArrayToxml.documentElement.appendChild(Node)
    Next
End Function

' 去除HTML标记
Public Function ReplaceHTML(Textstr)
    Dim Str, re
    Str = Textstr
    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = "<(.[^>]*)>"
    Str = re.Replace(Str, "")
    Set Re = Nothing
    ReplaceHTML = Str
End Function

Public Function MakeRandom(ByVal maxLen)
	Dim strNewPass,whatsNext, upper, lower, intCounter
	Randomize
	For intCounter = 1 To maxLen
		upper = 57:lower = 48:strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
	Next
	MakeRandom = strNewPass
End Function


Public Function ChkClng(ByVal str)
	If str<>"" and IsNumeric(str) Then
		ChkClng = CLng(str)
	Else
		ChkClng = 0
	End If
End Function

Public Function ChkCBool(ByVal str)
	If Not IsNull(str) Then
		ChkCBool = CBool(str)
	Else
		ChkCBool = False
	End If
End Function

Public Function ChkCDbl(ByVal str)
	If str<>"" and IsNumeric(str) Then
		ChkCDbl = CDbl(str)
	Else
		ChkCDbl = 0
	End If
End Function

Public Function ChkNull(ByVal str)
	If IsNull(str) Then
		ChkNull = ""
	Else
		ChkNull = str
	End If
End Function

' 判断是否安全字符串,在注册登录等特殊字段中使用
Public Function IsSafeStr(str)
	Dim s_BadStr, n, i
	s_BadStr = "' 　&<>?%,;:()`~!@#$^*{}[]|+-=" & Chr(34) & Chr(9) & Chr(32)
	n = Len(s_BadStr)
	IsSafeStr = True
	For i = 1 To n
		If Instr(str, Mid(s_BadStr, i, 1)) > 0 Then
			IsSafeStr = False
			Exit Function
		End If
	Next
End Function

' 全角半角转换函数
' flag=-1时进行半角转全角
' flag=0时进行半角全角互转
' flag=1时进行全角转半角
Function DBC2SBC(Str, flag)
    Dim i, sStr
    If Len(Str)<= 0 Then Exit Function
    DBC2SBC = ""
    For i = 1 To Len(Str)
        sStr = Asc(Mid(Str, i, 1))
        Select Case flag
            Case -1
                If sStr>0 And sStr<= 125 Then
                    DBC2SBC = DBC2SBC & Chr(Asc(Mid(Str, i, 1)) -23680)
                Else
                    DBC2SBC = DBC2SBC & Mid(Str, i, 1)
                End If
            Case 0
                If sStr>0 And sStr<= 125 Then
                    DBC2SBC = DBC2SBC & Chr(Asc(Mid(Str, i, 1)) -23680)
                Else
                    DBC2SBC = DBC2SBC & Chr(Asc(Mid(Str, i, 1)) + 23680)
                End If
            Case 1
                If sStr<0 Or sStr>125 Then
                    DBC2SBC = DBC2SBC & Chr(Asc(Mid(Str, i, 1)) + 23680)
                Else
                    DBC2SBC = DBC2SBC & Mid(Str, i, 1)
                End If
        End Select
    Next
End Function

' ****************************************
'	设置Cookie
' ****************************************
Public Function SetCookie(key, val, exptime)
	Response.Cookies(key) = val
	Response.Cookies(key).Expires = exptime
End Function

' ****************************************
'	读取Cookie
' ****************************************
Public Function GetCookie(key)
	GetCookie = Request.Cookies(key)
End Function


'Create a random code
	' IN  :
	' OUT :
	Function CreateRandomNumber(upperbound, lowerbound)
		Randomize
		CreateRandomNumber = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
	End Function


Function Server_VirtualPath(path)
    Dim RealPath, CurPath, i, Count, VirtualPath
    
    VirtualPath = Replace(path, chr(92), "/") & "/"
    RealPath = Split(VirtualPath, "/")
    CurPath = Split(Server.MapPath("./"), chr(92))
    Count = UBound(CurPath)
    
    For i = 0 To Count
        If CurPath(i) = RealPath(i) Then
            VirtualPath = Replace(VirtualPath, CurPath(i) & "/", "", 1, 1)
        Else
            VirtualPath = "../" & VirtualPath
        End If
    Next
    Server_Virtualpath = VirtualPath
End Function