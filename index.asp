<%@ Language="VBScript" CodePage=65001%>
<%
	Session.CodePage = 65001 
	Response.Charset = "UTF-8" 
	Response.Buffer = True 
	Server.ScriptTimeout = 99999	
	t1 = timer()		' 开始时间
	SQL_QUERY_NUM = 0		' 记录访问了多少次数据库
	Application_PATH = ""


	Dim c, m
	c = "home"
	m = "index"	

	' ******************************************
	' * 获取完全url地址
	' * @return				字符串
	' ******************************************
	Public Function GetUrl2() 
		Dim strTemp 
		If LCase(Request.ServerVariables("HTTPS")) = "off" Then 
			strTemp = "http://" 
		Else 
			strTemp = "https://" 
		End If 
		strTemp = strTemp & Request.ServerVariables("SERVER_NAME")		
		If Request.ServerVariables("SERVER_PORT") <> 80 Then strTemp = strTemp & ":" & Request.ServerVariables("SERVER_PORT") 
		strTemp = strTemp & Request.ServerVariables("URL")
		if Request.QueryString<> "" then
			strTemp = strTemp & "?" & Request.QueryString
		end if		
		GetUrl2 = strTemp
	End Function


	
	Application_PATH = Request.ServerVariables("http_host")	
	If GetUrl2() = "http://" & Request.ServerVariables("http_host") & "/index.asp" Then
		Dim filepath
		filepath = Request.ServerVariables("http_host") & ".htm"
		filepath = Server.MapPath(".") & "\" & filepath
		Dim FSO
		Set FSO = Server.CreateObject("Scripting.FileSystemObject")
		If FSO.FileExists(filepath) Then			
			Server.Transfer Request.ServerVariables("http_host") & ".htm"
		End If
		c = "manage"
		m = "index"
	End If		
	

	' ********************************************
	' * 两个Cache函数
	'
	' ********************************************
	Function GetCacheValue(cacheToken)

		If CDate(Application("exp_" & Application_PATH & cacheToken)) >= Now Then
			GetCacheValue = Application("data_" & Application_PATH & cacheToken)
		Else   
			GetCacheValue = ""
		End If

	End Function

	Function SetCacheValue(cacheToken, cacheValue, cacheSecond)
		Application.Lock    
		Application("data_" & Application_PATH & cacheToken) = cacheValue
		Application("exp_" & Application_PATH & cacheToken) = DateAdd("s", cacheSecond, Now)
		Application.Unlock 
	End Function

	' 仿PHP输出函数
	Function Echo(Str)
		Response.Write Str & VbCrlf
	End Function

	Sub Redirect(url)
		Response.Redirect("index.asp?/" & url )
	End Sub

	Function GetDomain()
		GetDomain = Request.ServerVariables("http_host")
	End Function
%>

<%	
	
	Sub Import(filepath)

		Dim retVal		
		filepath = Server.MapPath(".") & "\" & filepath
		'retVal = GetCacheValue(filepath)


		'If retVal = "" Then 	

			' ************************************************
			' 此段程序用来读取gb2312编码的文件
			'Const ForReading = 1	
			
			'Set objFSO = CreateObject("Scripting.FileSystemObject")
			'Set objTextFile = objFSO.OpenTextFile(filepath, ForReading)		
			'retVal = objTextFile.ReadAll()
			' *************************************************
			
			Set filestream = Server.CreateObject("ADODB.Stream")
			With filestream			
				.Type = 2 '以本模式读取
				.Mode = 3 
				.Charset = "utf-8"
				.Open
				.Loadfromfile filepath
				retVal = .readtext
				.Close
			End With
			Set filestream = Nothing
		



			'SetCacheValue filepath, retVal, 180

		'End If 		

		ExecuteGlobal retVal	
			
	End Sub	

	Function LoadModel(model)		
		Import "system/" & Application_PATH & "/models/" & model & ".vbs"	
		Execute "Set LoadModel = New " & model		 
	End Function

%>


<%
	
	' 自定装载类，装载后可在控制器中直接实例化类或调用函数不用	
	Import "system/" & Application_PATH & "/config/config.vbs"
	Import "system/libraries/StringBuilder.vbs"	
	Import "system/libraries/clsTemplate.vbs"	
	Import "system/libraries/clsPager.vbs"
	Import "system/libraries/JSON.vbs"
	

	' 装载全局函数
	Import "system/helpers/md5.vbs"
	'Import "system/helpers/sha1.vbs"
	Import "system/helpers/file_helper.vbs"
	Import "system/helpers/string_helper.vbs"
	Import "system/helpers/util_helper.vbs"
	Import "system/" & Application_PATH & "/plugin/plugins.vbs"
	Import "system/helpers/date_helper.vbs"


	Import "system/libraries/clsDB.vbs"		
	Set db = New DBControl
	db.Open

	Import "system/libraries/clsTagParser.vbs"
	Dim p
	Set p = New TagParser

	Dim t
	Set t = New ccClsTemplate

	Dim sb
	Set sb = New StringBuilder

	'开启缓存
	't.PublicCache = True
	't.PrivateCache = True
	
	
	
	
	If (Request.QueryString("c") <> "") Then
		c = Request.QueryString("c")
	End If
	If (Request.QueryString("m") <> "") Then
		m = Request.QueryString("m")
	End If

	Dim segment(100)

	For i = 0 To UBound(segment)
		segment(i) = "0"
	Next

	Dim tempSegment

	If Request.ServerVariables("QUERY_STRING") <> "" Then
	
	
		tempSegment = Split(Request.ServerVariables("QUERY_STRING"), "/")

		For i = 0 To UBound(tempSegment)
			segment(i) = tempSegment(i)
		Next
		
		If config("pathinfo") = 1 Then
			'On Error Resume Next
			c = tempSegment(1)
			m = tempSegment(2)	
		End If
	End If
	

	Dim d
	Set d = CreateObject("Scripting.Dictionary")
	

	Import "system/" & Application_PATH & "/controllers/" & c & ".vbs"

	Execute "Set objController = New " & c
	Execute "Call objController." & m



	Set d = Nothing
	Set t = Nothing	

	Response.Write p.Parser(sb.ToString())

	Set p = Nothing

	db.Close
	Set db = Nothing

	Set sb = Nothing

	If config("isdebug") = 1 Then
		Response.Write "<br /><div style=""font-size: 12px;"">执行时间：" & FormatNumber(timer() - t1, 5, -1) & "毫秒。查询数据库" & SQL_QUERY_NUM & "次。</div>" & vbCrLf
	End If

	Set config = Nothing	
%>



