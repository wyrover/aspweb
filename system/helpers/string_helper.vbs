' **********************************
' * 字符串长度
' * @Str		字符串
' * @return	字符串长度
' **********************************
Public Function StrLength(Str)
    On Error Resume Next
    Dim WINNT_CHINESE
    WINNT_CHINESE = (Len("中国") = 2)
    If WINNT_CHINESE Then
        Dim l, t, c
        Dim i
        l = Len(Str)
        t = l
        For i = 1 To l
            c = Asc(Mid(Str, i, 1))
            If c < 0 Then c = c + 65536
            If c > 255 Then
                t = t + 1
            End If
        Next
        StrLength = t
    Else
    StrLength = Len(Str)
    End If
    If Err.Number <> 0 Then Err.Clear
End Function


'*************************************
'过滤特殊字符
'*************************************
Function CheckStr(byVal ChkStr) 
	Dim Str:Str=ChkStr
	If IsNull(Str) Then
		CheckStr = ""
		Exit Function 
	End If
    Str = Replace(Str, "&", "&amp;")
    Str = Replace(Str,"'","&#39;")
    Str = Replace(Str,"""","&#34;")
	Dim re
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	re.Pattern="(w)(here)"
    Str = re.replace(Str,"$1h&#101;re")
	re.Pattern="(s)(elect)"
    Str = re.replace(Str,"$1el&#101;ct")
	re.Pattern="(i)(nsert)"
    Str = re.replace(Str,"$1ns&#101;rt")
	re.Pattern="(c)(reate)"
    Str = re.replace(Str,"$1r&#101;ate")
	re.Pattern="(d)(rop)"
    Str = re.replace(Str,"$1ro&#112;")
	re.Pattern="(a)(lter)"
    Str = re.replace(Str,"$1lt&#101;r")
	re.Pattern="(d)(elete)"
    Str = re.replace(Str,"$1el&#101;te")
	re.Pattern="(u)(pdate)"
    Str = re.replace(Str,"$1p&#100;ate")
	re.Pattern="(\s)(or)"
    Str = re.replace(Str,"$1o&#114;")
	Set re=Nothing
	CheckStr=Str
End Function

'*************************************
'恢复特殊字符
'*************************************
Function UnCheckStr(ByVal Str)
		If IsNull(Str) Then
			UnCheckStr = ""
			Exit Function 
		End If
	    Str = Replace(Str,"&#39;","'")
        Str = Replace(Str,"&#34;","""")
		Dim re
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		re.Pattern="(w)(h&#101;re)"
	    str = re.replace(str,"$1here")
		re.Pattern="(s)(el&#101;ct)"
	    str = re.replace(str,"$1elect")
		re.Pattern="(i)(ns&#101;rt)"
	    str = re.replace(str,"$1nsert")
		re.Pattern="(c)(r&#101;ate)"
	    str = re.replace(str,"$1reate")
		re.Pattern="(d)(ro&#112;)"
	    str = re.replace(str,"$1rop")
		re.Pattern="(a)(lt&#101;r)"
	    str = re.replace(str,"$1lter")
		re.Pattern="(d)(el&#101;te)"
	    str = re.replace(str,"$1elete")
		re.Pattern="(u)(p&#100;ate)"
	    str = re.replace(str,"$1pdate")
		re.Pattern="(\s)(o&#114;)"
	    Str = re.replace(Str,"$1or")
		Set re=Nothing
        Str = Replace(Str, "&amp;", "&")
    	UnCheckStr=Str
End Function


' **************************************
' * 获取若干位随机字符串
' * @digits			字符位数
' * @return			字符串
' **************************************
Function GetRandomString(digits)
'定义并初始化数组
    Dim char_array(80)
    '初始化数字
    For i = 0 To 9
        char_array(i) = CStr(i)
    Next
    '初始化大写字母
    For i = 10 To 35
        char_array(i) = Chr(i + 55)
    Next
    '初始化小写字母
    For i = 36 To 61
        char_array(i) = Chr(i + 61)
    Next
    Randomize   '初始化随机数生成器。
    Do While Len(output) < digits
        num = char_array(Int((62 - 0 + 1) * Rnd + 0))
        output = output + num
    Loop
'设置返回值
    gen_key = output
End Function


' *******************************
' *	 危险SQL过滤
' * @strValue	SQL
' * @return		过滤结果
' *******************************
Function CleanForSQL(strValue)
	'create another copy
	Dim strTemp
	strTemp = strValue

	'clean out single quotes
	Do Until InStr(1, strTemp, "'") = 0
	  strTemp = Left(strTemp, InStr(1, strTemp, "'") - 1) & Right(strTemp, Len(strTemp) - InStr(1, strTemp, "'"))
	Loop
	   
	'return the clean string
	CleanForSQL = strTemp
End Function


' *******************************
' * 过滤html标记
' * @str			字符串
' * @return		字符串
' *******************************
Function FilterHtml(str)
	Dim regEx
	'创建正则对象
	Set regEx = New RegExp
	regEx.IgnoreCase = True
	regEx.Global = True
	regEx.MultiLine = True

	regEx.Pattern = "<.+?>"
	FilterHtml = regEx.Replace(str,"")
	Set regEx = Nothing
End Function


' *******************************
' * html编码
' * @str			字符串
' * @return		字符串
' *******************************
Function HTMLEncode(Byval fString)
	If Not IsNull(fString) then
	fString = replace(fString, ">", "&gt;")
	fString = replace(fString, "<", "&lt;")
	fString = replace(fString, "&", "&amp;")
	fString = Replace(fString, CHR(32), "&nbsp;")
	fString = Replace(fString, CHR(9), "&nbsp;")
	fString = Replace(fString, CHR(34), "&quot;")
	fString = Replace(fString, CHR(39), "&#39;")
	fString = Replace(fString, CHR(13), "")
	fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
	fString = Replace(fString, CHR(10), "<br /> ")
	HTMLEncode = fString
	End If
End Function

Public Function HTMLDeCode(Byval fString)
	If Not IsNull(fString) then
	fString = replace(fString, "&gt;", ">")
	fString = replace(fString, "&lt;", "<")
	fString = Replace(fString,  "&nbsp;"," ")
	fString = Replace(fString, "&quot;", CHR(34))
	fString = Replace(fString, "&#39;", CHR(39))
	fString = Replace(fString, "</P><P> ",CHR(10) & CHR(10))
	fString = Replace(fString, "<br /> ", CHR(10))
	HTMLDeCode = fString
	End If
End Function


' *******************************
' * 字符串截断(可识别中英文)
' * @str			预截字符串
' * @strlen		长度
' * @return		字符串
' *******************************
Function CutStr(str,strlen)
	DIM l,t,c,m_i
	l=len(str)
	t=0
	For m_i = 1 To l
		c = Abs(Asc(Mid(str,m_i,1)))
		If c > 255 Then
			t = t+2
		Else
			t = t+1
		End If

		If t >= strlen Then
			CutStr = left(str,m_i) & "..."
			exit for
		Else
			CutStr = str
		End if
	Next
End Function 

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

  Str=Str_unicode
end function


Public Function DecodeFilter(html, filter)
	html=LCase(html)
	filter=split(filter,",")
	For Each i In filter
		Select Case i
			Case "SCRIPT"		' 去除所有客户端脚本javascipt,vbscript,jscript,js,vbs,event,...
				html = exeRE("(javascript|jscript|vbscript|vbs):", "#", html)
				html = exeRE("</?script[^>]*>", "", html)
				html = exeRE("on(mouse|exit|error|click|key)", "", html)
			Case "TABLE":		' 去除表格<table><tr><td><th>
				html = exeRE("</?table[^>]*>", "", html)
				html = exeRE("</?tr[^>]*>", "", html)
				html = exeRE("</?th[^>]*>", "", html)
				html = exeRE("</?td[^>]*>", "", html)
				html = exeRE("</?tbody[^>]*>", "", html)
			Case "CLASS"		' 去除样式类class=""
				html = exeRE("(<[^>]+) class=[^ |^>]*([^>]*>)", "$1 $2", html) 
			Case "STYLE"		' 去除样式style=""
				html = exeRE("(<[^>]+) style=""[^""]*""([^>]*>)", "$1 $2", html)
				html = exeRE("(<[^>]+) style='[^']*'([^>]*>)", "$1 $2", html)
			Case "IMG"		' 去除样式style=""
				html = exeRE("</?img[^>]*>", "", html)
			Case "XML"		' 去除XML<?xml>
				html = exeRE("<\\?xml[^>]*>", "", html)
			Case "NAMESPACE"	' 去除命名空间<o:p></o:p>
				html = exeRE("<\/?[a-z]+:[^>]*>", "", html)
			Case "FONT"		' 去除字体<font></font>
				html = exeRE("</?font[^>]*>", "", html)
			Case "MARQUEE"		' 去除字幕<marquee></marquee>
				html = exeRE("</?marquee[^>]*>", "", html)
			Case "OBJECT"		' 去除对象<object><param><embed></object>
				html = exeRE("</?object[^>]*>", "", html)
				html = exeRE("</?param[^>]*>", "", html)
				html = exeRE("</?embed[^>]*>", "", html)
			Case "DIV"		' 去除对象<object><param><embed></object>
				html = exeRE("</?div([^>])*>", "$1", html)
		End Select
	Next
	'html = Replace(html,"<table","<")
	'html = Replace(html,"<tr","<")
	'html = Replace(html,"<td","<")
	DecodeFilter = html
End Function

Public Function exeRE(re, rp, content)
	Set oReg = New RegExp
	oReg.IgnoreCase =True
	oReg.Global=True	
	oReg.Pattern=re
	r = oReg.Replace(content,rp)
	Set oReg = Nothing	
	exeRE = r
End Function


Function Test(a)
	Test = "<img src=" & a & ">"
End Function

Public Function MsgBox2(HintText,HintType,GoWhere)
	Dim Hint,HintTypeText
	Select Case HintType
		Case "0"
			Hint=16
			HintTypeText="出错啦！"
		Case "1" 
			Hint=48
			HintTypeText="警告!"
		Case "2" 
			Hint=64
			HintTypeText="提示!"
	End Select
	Response.Write "<Script Language=VBScript>"
	Response.Write "MsgBox """ & Replace(HintText,"'","") &_
		"""," & Hint & ",""" & HintTypeText & """ "
	Response.Write "</Script>"
	if GoWhere<>"" then
		if GoWhere = "0" then
			Response.Write "<Script Language=JavaScript>history.back();</Script>"
		else
			Response.Write "<Script Language=JavaScript>location.href='" & GoWhere & "';</Script>"
		end if
	end if
	Response.End()
End Function

'=============================================================
	'函数名:ChkFormStr
	'作  用:过滤表单字符
	'参  数:str   ----原字符串
	'返回值:过滤后的字符串
	'=============================================================
	Public Function ChkFormStr(ByVal str)
		Dim fString
		fString = str
		If IsNull(fString) Then
			ChkFormStr = ""
			Exit Function
		End If
		fString = Replace(fString, "'", "&#39;")
		fString = Replace(fString, Chr(34), "&quot;")
		fString = Replace(fString, Chr(13), "")
		fString = Replace(fString, Chr(10), "")
		fString = Replace(fString, Chr(9), "")
		fString = Replace(fString, ">", "&gt;")
		fString = Replace(fString, "<", "&lt;")
		fString = Replace(fString, "%", "%")
		ChkFormStr = Trim(JAPEncode(fString))
	End Function
	'=============================================================
	'函数作用:过滤SQL非法字符
	'=============================================================
	Public Function CheckRequest(ByVal str,ByVal strLen)
		On Error Resume Next
		str = Trim(str)
		str = Replace(str, Chr(0), "")
		str = Replace(str, "'", "")
		str = Replace(str, "%", "")
		str = Replace(str, "^", "")
		str = Replace(str, ";", "")
		str = Replace(str, "*", "")
		str = Replace(str, "<", "")
		str = Replace(str, ">", "")
		str = Replace(str, "|", "")
		str = Replace(str, "and", "")
		str = Replace(str, "chr", "")
		
		If Len(str) > 0 And strLen > 0 Then
			str = Left(str, strLen)
		End If
		CheckRequest = str
	End Function


Function StripHTML(ByRef sHTML)
	Dim re	' Regular Expression Object

	' Create Regular Expression object
	Set re = New RegExp
	re.Pattern = "<[^>]*>" ' Set the pattern to look For "<anychar>" tags
	re.IgnoreCase = True   ' Set case insensitivity.
	re.Global = True       ' Set global applicability.

	StripHTML = re.Replace(sHTML, " ") ' Return the original String stripped of HTML

	' Release object from memory
	Set re = Nothing
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
	'函数名：ChkBadWords
	'作  用：屏蔽字符
	'参  数：str-----要屏蔽的字符
	'返回值：替换屏蔽后的字符
	'**************************************************
	Public Function ChkBadWords(Str)
		
		ChkBadWords = Str
	End Function

	'**************************************************
	'函 数 名：CID
	'作    用：转换为有效的 ID
	'返回值类型：Integer (>=0)
	'**************************************************
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


'*************************************
'过滤特殊字符
'*************************************
Function CheckStr(byVal ChkStr) 
	Dim Str:Str=ChkStr
	If IsNull(Str) Then
		CheckStr = ""
		Exit Function 
	End If
    Str = Replace(Str, "&", "&amp;")
    Str = Replace(Str,"'","&#39;")
    Str = Replace(Str,"""","&#34;")
	Dim re
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	re.Pattern="(w)(here)"
    Str = re.replace(Str,"$1h&#101;re")
	re.Pattern="(s)(elect)"
    Str = re.replace(Str,"$1el&#101;ct")
	re.Pattern="(i)(nsert)"
    Str = re.replace(Str,"$1ns&#101;rt")
	re.Pattern="(c)(reate)"
    Str = re.replace(Str,"$1r&#101;ate")
	re.Pattern="(d)(rop)"
    Str = re.replace(Str,"$1ro&#112;")
	re.Pattern="(a)(lter)"
    Str = re.replace(Str,"$1lt&#101;r")
	re.Pattern="(d)(elete)"
    Str = re.replace(Str,"$1el&#101;te")
	re.Pattern="(u)(pdate)"
    Str = re.replace(Str,"$1p&#100;ate")
	re.Pattern="(\s)(or)"
    Str = re.replace(Str,"$1o&#114;")
	Set re=Nothing
	CheckStr=Str
End Function

'*************************************
'恢复特殊字符
'*************************************
Function UnCheckStr(ByVal Str)
		If IsNull(Str) Then
			UnCheckStr = ""
			Exit Function 
		End If
	    Str = Replace(Str,"&#39;","'")
        Str = Replace(Str,"&#34;","""")
		Dim re
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		re.Pattern="(w)(h&#101;re)"
	    str = re.replace(str,"$1here")
		re.Pattern="(s)(el&#101;ct)"
	    str = re.replace(str,"$1elect")
		re.Pattern="(i)(ns&#101;rt)"
	    str = re.replace(str,"$1nsert")
		re.Pattern="(c)(r&#101;ate)"
	    str = re.replace(str,"$1reate")
		re.Pattern="(d)(ro&#112;)"
	    str = re.replace(str,"$1rop")
		re.Pattern="(a)(lt&#101;r)"
	    str = re.replace(str,"$1lter")
		re.Pattern="(d)(el&#101;te)"
	    str = re.replace(str,"$1elete")
		re.Pattern="(u)(p&#100;ate)"
	    str = re.replace(str,"$1pdate")
		re.Pattern="(\s)(o&#114;)"
	    Str = re.replace(Str,"$1or")
		Set re=Nothing
        Str = Replace(Str, "&amp;", "&")
    	UnCheckStr=Str
End Function


'*************************************
'切割内容 - 按行分割
'*************************************
Function SplitLines(byVal Content,byVal ContentNums) 
	Dim ts,i,l
	ContentNums=int(ContentNums)
	If IsNull(Content) Then Exit Function
	i=1
	ts = 0
	For i=1 to Len(Content)
      l=Lcase(Mid(Content,i,5))
      	If l="<br/>" Then
         	ts=ts+1
      	End If
      l=Lcase(Mid(Content,i,4))
      	If l="<br>" Then
         	ts=ts+1
      	End If
      l=Lcase(Mid(Content,i,3))
      	If l="<p>" Then
         	ts=ts+1
      	End If
	If ts>ContentNums Then Exit For 
	Next
	If ts>ContentNums Then
    	Content=Left(Content,i-1)
	End If
	SplitLines=Content
End Function


Public Function DecodeFilter(Byval sContent, Byval sFilter)
	Dim regEx
	Set regEx = New RegExp
	regEx.IgnoreCase = True
	regEx.Global	 = True
	Select Case Ucase(sFilter)
	Case "SCRIPT"'去除所有客户端脚本javascipt,vbscript,jscript,js,vbs,event,...
		regEx.Pattern	= "</?script[^>]*>"
		sContent		= regEx.replace(sContent,"")
		regEx.Pattern	= "(javascript|jscript|vbscript|vbs):"
		sContent		= regEx.replace(sContent,"$1：")
		'regEx.Pattern	= "on(mouse|exit|error|click|key)"
		'sContent		= regEx.replace(sContent,"<I>on$1</I>")
	Case "OBJECT"'去除对象<object><param><embed></object>
		regEx.Pattern	= "</?object[^>]*>"
		sContent		= regEx.replace(sContent,"")
		regEx.Pattern	= "</?param[^>]*>"
		sContent		= regEx.replace(sContent,"")
		regEx.Pattern	= "</?embed[^>]*>"
		sContent		= regEx.replace(sContent,"")
	Case "TABLE"'去除表格<table><tr><td><th>
		regEx.Pattern	= "</?table[^>]*>"
		sContent		= regEx.replace(sContent,"")
		regEx.Pattern	= "</?tr[^>]*>"
		sContent		= regEx.replace(sContent,"")
		regEx.Pattern	= "</?th[^>]*>"
		sContent		= regEx.replace(sContent,"")
		regEx.Pattern	= "</?td[^>]*>"
		sContent		= regEx.replace(sContent,"")
	Case "CLASS"'去除样式类class=""
		regEx.Pattern	= "(<[^>]+) class=[^ |^>]*([^>]*>)"
		sContent		= regEx.replace(sContent,"$1 $2")
	Case "STYLE"'去除样式style=""
		regEx.Pattern	= "(<[^>]+) style=\""[^\""]*\""([^>]*>)"
		sContent		= regEx.replace(sContent,"")
	Case "XML"'去除XML<?xml>
		regEx.Pattern	= "<\\?xml[^>]*>"
		sContent		= regEx.replace(sContent,"")
	Case "NAMESPACE"'去除命名空间<o:p></o:p>
		regEx.Pattern	= "<\/?[a-z]+:[^>]*>"
		sContent		= regEx.replace(sContent,"")
	Case Else
		regEx.Pattern	= "</?" & sFilter & "[^>]*>"
		sContent		= regEx.replace(sContent,"")
	End Select
	DecodeFilter = sContent
	Set regEx=nothing
End Function


'sContent（要转换的数据字符串）
'sFilters（要过滤掉的格式集，用"|"分隔多个）
Public Function DeCode(Byval sContent, Byval sFilters)
	Dim a_Filter, i, s_Result, s_Filters
	Decode = sContent
	If IsNull(sContent) or IsNull(sFilters) Then Exit Function
	If sContent = "" or sFilters = "" Then Exit Function
	s_Result  = sContent
	s_Filters = sFilters
	If InStr(s_Filters,"|")>0 then
		a_Filter = Split(s_Filters, "|")
		For i = 0 To UBound(a_Filter)
			s_Result = DecodeFilter(s_Result, a_Filter(i))
		Next
	Else
		s_Result = DecodeFilter(s_Result, s_Filters)
	End If
	DeCode = s_Result
End Function

Public Function NoHtml(Byval str)
	if not isnull(str) then
	dim re
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	re.Pattern="(\<.[^\<]*\>)"
	str=re.replace(str," ")
	re.Pattern="(\<\/[^\<]*\>)"
	str=re.replace(str," ")
	NoHtml=str
	Set re=Nothing
	End if
End Function

'检查Email地址合法性
Public Function ChkEmail(email)
	dim names, name, i, c
	ChkEmail = true : names = Split(email, "@")
	if UBound(names) <> 1 then ChkEmail = false : Exit Function
	for each name in names
		if Len(name) <= 0 then ChkEmail = false:exit function
		for i = 1 to Len(name)
			c = Lcase(Mid(name, i, 1))
			if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
			   ChkEmail = false:Exit Function
			end if
		next
		if Left(name, 1) = "." or Right(name, 1) = "." then
			ChkEmail = false:Exit Function
		end if
	next
	if InStr(names(1), ".") <= 0 then ChkEmail = false:exit function
	i = Len(names(1)) - InStrRev(names(1), ".")
	if i <> 2 and i <> 3 then ChkEmail = false : Exit function
	if InStr(email, "..") > 0 then ChkEmail = false
End Function





'向地址中加入 ? 或 & 
Public Function JoinChar(strUrl)
	if strUrl="" then JoinChar="":Exit Function
	if InStr(strUrl,"?")<len(strUrl) then 
		if InStr(strUrl,"?")>1 then
			if InStr(strUrl,"&")<len(strUrl) then 
				JoinChar=strUrl & "&"
			else
				JoinChar=strUrl
			end if
		else
			JoinChar=strUrl & "?"
		end if
	else
		JoinChar=strUrl
	end if
End Function

Public Function GetTitleFont(Byval sValue,Byval sType)
	Select Case ChkClng(sType)
	Case 0 : GetTitleFont = sValue
	Case 1 : GetTitleFont = "<strong>" & sValue & "</strong>"
	Case 2 : GetTitleFont = "<em>" & sValue & "</em>"
	Case 3 : GetTitleFont = "<strong><em>" & sValue & "</em></strong>"
	Case 4 : GetTitleFont = "<u>" & sValue & "</u>"
	Case 5 : GetTitleFont = "<strong><u>" & sValue & "</u></strong>"
	Case 6 : GetTitleFont = "<em><u>" & sValue & "</u></em>"
	Case 7 : GetTitleFont = "<strong><em><u>" & sValue & "</u></em></strong>"
	Case Else : GetTitleFont = sValue
	End Select
End Function

Public Function FormatColor(Byval sValue,Byval sColor)
	sColor=Trim(sColor)
	if Isnull(sColor) or sColor="" then FormatColor=sValue : Exit Function
	FormatColor = "<font color="""& sColor &""">" & sValue & "</font>"
End Function

'过滤非法的SQL字符
Public Function ReplaceBadChar(Byval strChar)
	strChar=replace(replace(strChar," ",""),"'","")
	strChar=replace(replace(strChar,".",""),"<","")
	strChar=replace(replace(strChar,")",""),"(","")
	strChar=replace(replace(strChar,"?",""),"*","")
	strChar=replace(replace(strChar,"/",""),"\","")
	ReplaceBadChar=replace(strChar,Chr(0),"")
End Function


Function CheckStr(byVal s)
	s = Trim(s)
	
	If IsNull(s) Then
		CheckStr = ""
		Exit Function 
	End If
	
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "'", "&#39;")
    s = Replace(s, """", "&#34;")
	
	'    正则，替换 SQL 关键词
	Dim re
	Set re = New RegExp
	re.IgnoreCase = True
	re.Global = True
	re.Pattern = "(w)(here)"
    s = re.Replace(s, "$1h&#101;re")
	re.Pattern = "(s)(elect)"
    s = re.Replace(s, "$1el&#101;ct")
	re.Pattern = "(i)(nsert)"
    s = re.Replace(s,  "$1ns&#101;rt")
	re.Pattern = "(c)(reate)"
    s = re.Replace(s,"$1r&#101;ate")
	re.Pattern = "(d)(rop)"
    s = re.Replace(s, "$1ro&#112;")
	re.Pattern = "(a)(lter)"
    s = re.Replace(s, "$1lt&#101;r")
	re.Pattern = "(d)(elete)"
    s = re.Replace(s, "$1el&#101;te")
	re.Pattern = "(u)(pdate)"
    s = re.Replace(s, "$1p&#100;ate")
	re.Pattern = "(\s)(or)"
    s = re.Replace(s, "$1o&#114;")
	Set re = Nothing
	CheckStr = s
End Function




Function UnCheckStr(ByVal Str)
	If IsNull(Str) Then
		UnCheckStr = ""
		Exit Function 
	End If
	    Str = Replace(Str,"&#39;","'")
        Str = Replace(Str,"&#34;","""")
		Dim re
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		re.Pattern="(w)(h&#101;re)"
	    str = re.replace(str,"$1here")
		re.Pattern="(s)(el&#101;ct)"
	    str = re.replace(str,"$1elect")
		re.Pattern="(i)(ns&#101;rt)"
	    str = re.replace(str,"$1nsert")
		re.Pattern="(c)(r&#101;ate)"
	    str = re.replace(str,"$1reate")
		re.Pattern="(d)(ro&#112;)"
	    str = re.replace(str,"$1rop")
		re.Pattern="(a)(lt&#101;r)"
	    str = re.replace(str,"$1lter")
		re.Pattern="(d)(el&#101;te)"
	    str = re.replace(str,"$1elete")
		re.Pattern="(u)(p&#100;ate)"
	    str = re.replace(str,"$1pdate")
		re.Pattern="(\s)(o&#114;)"
	    Str = re.replace(Str,"$1or")
		Set re=Nothing
        Str = Replace(Str, "&amp;", "&")
    	UnCheckStr=Str
End Function


function filterinput(str)
	dim regex   
	set regex = new regexp         
	regex.pattern = "[^\w\d]"       
	regex.ignorecase = true        
	regex.global = true         
	regex.multiline = true
	filterinput = regex.replace(str,"")
end function

'*************************************
'自动闭合HTML
'*************************************
function closeHTML(strContent)
  dim arrTags,i,OpenPos,ClosePos,re,strMatchs,j,Match
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
    arrTags=array("p","div","span","table","ul","font","b","u","i","h1","h2","h3","h4","h5","h6")
  for i=0 to ubound(arrTags)
   OpenPos=0
   ClosePos=0
   
   re.Pattern="\<"+arrTags(i)+"( [^\<\>]+|)\>"
   Set strMatchs=re.Execute(strContent)
   For Each Match in strMatchs
    OpenPos=OpenPos+1
   next
   re.Pattern="\</"+arrTags(i)+"\>"
   Set strMatchs=re.Execute(strContent)
   For Each Match in strMatchs
    ClosePos=ClosePos+1
   next
   for j=1 to OpenPos-ClosePos
      strContent=strContent+"</"+arrTags(i)+">"
   next
  next
closeHTML=strContent
end function



function URLDecode(enStr)
  dim  deStr,strSpecial
  dim  c,i,v
  deStr=""
  strSpecial="!""#$%&'()*+,/:;<=>?@[\]^`{ |}~%"
  for  i=1  to  len(enStr)
    c=Mid(enStr,i,1)
    if  c="%"  then
    v=eval("&h"+Mid(enStr,i+1,2))
    if  inStr(strSpecial,chr(v))>0  then
    deStr=deStr&chr(v)
    i=i+2
    else
    v=eval("&h"+Mid(enStr,i+1,2)+Mid(enStr,i+4,2))
    deStr=deStr&chr(v)
    i=i+5
    end  if
    else
    if  c="+"  then
    deStr=deStr&" "
    else
    deStr=deStr&c
    end  if
    end  if
  next
  URLDecode=deStr
end function