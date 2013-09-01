
Class clsThief


    Private value_ ' 窃取到的内容
    Private src_ ' 要偷的目标URL地址
    Private isGet_ ' 判断是否已经偷过
	Private cookie_ 

    ' 赋值—要偷的目标URL地址/属性

    Public Property Let src(Str)
        src_ = Str
    End Property

    '返回值—最终窃取并应用类方法加工过的内容/属性

    Public Property Get Value
        Value = value_
    End Property

	Public Property Get Cookie
        Cookie = cookie_
    End Property

    Public Property Get Version
        Version = "先锋海盗类 Version 2004"
    End Property

    Private Sub class_initialize()
        value_ = ""
        src_ = ""
        isGet_ = False
    End Sub

    Private Sub class_terminate()
    End Sub

    ' 中文处理

    Private Function BytesToBstr(body, Cset)
        Dim objstream
        Set objstream = CreateObject("adodb.stream")
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

    ' 窃取目标URL地址的HTML代码/方法

    Public Sub steal(encode)
        If src_<>"" Then
            Dim Http
            Set Http = CreateObject("MSXML2.XMLHTTP")
            Http.Open "GET", src_ , false
            Http.send()
			'cookie = Http.getResponseHeader("Set-Cookie")
            If Http.readystate<>4 Then
                Exit Sub
            End If
            value_ = BytesToBSTR(Http.responseBody, encode)
            isGet_ = True
            Set http = Nothing
            If Err.Number<>0 Then Err.Clear
        Else
            response.Write("<script>alert(""请先设置src属性！"")</script>")
        End If
    End Sub

    ' 删除偷到的内容中里面的换行、回车符以便进一步加工/方法

    Public Sub noReturn()
        If isGet_ = false Then Call steal()
        value_ = Replace(Replace(value_ , vbCr, ""), vbLf, "")
    End Sub

    ' 对偷到的内容中的个别字符串用新值更换/方法
	' 参数分别是旧字符串,新字符串
    Public Sub change(oldStr, Str) 
        If isGet_ = false Then Call steal()
        value_ = Replace(value_ , oldStr, Str)
    End Sub

    ' 按指定首尾字符串对偷取的内容进行裁减（不包括首尾字符串）/方法
    ' 参数分别是首字符串,尾字符串

    Public Sub cut(head, bot)
        If isGet_ = false Then Call steal()
        value_ = Mid(value_ , InStr(value_ , head) + Len(head), InStr(value_ , bot) - InStr(value_ , head) - Len(head))
    End Sub

    ' 按指定首尾字符串对偷取的内容进行裁减（包括首尾字符串）/方法
    ' 参数分别是首字符串,尾字符串

    Public Sub cutX(head, bot)
        If isGet_ = false Then Call steal()
        value_ = Mid(value_ , InStr(value_ , head), InStr(value_ , bot) - InStr(value_ , head) + Len(bot))
    End Sub

    '按指定首尾字符串位置偏移指针对偷取的内容进行裁减/方法
    '参数分别是首字符串,首偏移值,尾字符串,尾偏移值,左偏移用负值,偏移指针单位为字符数

    Public Sub cutBy(head, headCusor, bot, botCusor)
        If isGet_ = false Then Call steal()
        value_ = Mid(value_ , InStr(value_ , head) + Len(head) + headCusor, InStr(value_ , bot) -1 + botCusor - InStr(value_ , head) - Len(head) - headcusor)
    End Sub

    '按指定首尾字符串对偷取的内容用新值进行替换（不包括首尾字符串）/方法
    '参数分别是首字符串,尾字符串,新值,新值位空则为过滤

    Public Sub filt(head, bot, Str)
        If isGet_ = false Then Call steal()
        value_ = Replace(value_, Mid(value_ , InStr(value_ , head) + Len(head), InStr(value_ , bot) -1), Str)
    End Sub

    '按指定首尾字符串对偷取的内容用新值进行替换（包括首尾字符串）/方法
    '参数分别是首字符串,尾字符串,新值,新值为空则为过滤

    Public Sub filtX(head, bot, Str)
        If isGet_ = false Then Call steal()
        value_ = Replace(value_, Mid(value_ , InStr(value_ , head), InStr(value_ , bot) + Len(bot) -1), Str)
    End Sub

    '按指定首尾字符串位置偏移指针对偷取的内容新值进行替换/方法
    '参数分别是首字符串,首偏移值,尾字符串,尾偏移值,新值,左偏移用负值,偏移指针单位为字符数,新值为空则为过滤

    Public Sub filtBy(head, headCusor, bot, botCusor, Str)

        If isGet_ = false Then Call steal()
        value_ = Replace(value_ , Mid(value_ , InStr(value_ , head) + Len(head) + headCusor, InStr(value_ , bot) -1 + botCusor - InStr(value_ , head) - Len(head) - headcusor), Str)
    End Sub

    '将偷取的内容中的绝对URL地址改为本地相对地址

    Public Sub local()
        Dim tempReg
        Set tempReg = New RegExp
        tempReg.IgnoreCase = true
        tempReg.Global = true
        tempReg.Pattern = "^(http|https|ftp):(\/\/|////)(\w+.)+(com|net|org|cc|tv|cn|biz|com.cn|net.cn|sh.cn)\/"
        value_ = tempReg.Replace(value_ , "")
        Set tempReg = Nothing
    End Sub

    '对偷到的内容中的符合正则表达式的字符串用新值进行替换/方法
    '参数是你自定义的正则表达式,新值

    Public Sub replaceByReg(patrn, Str)
        If isGet_ = false Then Call steal()
        Dim tempReg
        Set tempReg = New RegExp
        tempReg.IgnoreCase = true
        tempReg.Global = true
        tempReg = patrn
        value_ = tempReg.Replace(value_ , Str)
        Set tempReg = Nothing
    End Sub

    '应用正则表达式对符合条件的内容进行分块采集并组合,最终内容为以<!--lkstar-->隔断的大文本/方法
    '通过属性value得到此内容后你可以用split(value,"<!--lkstar-->")得到你需要的数组
    '参数是你自定义的正则表达式

    Public Sub pickByReg(patrn)
        If isGet_ = false Then Call steal()
        Dim tempReg, match, matches, content
        Set tempReg = New RegExp
        tempReg.IgnoreCase = true
        tempReg.Global = true
        tempReg = patrn
        Set matches = tempReg.Execute(value_)
        For Each match in matches
            content = content&match.Value&"<!--lkstar-->"
        Next
        value_ = content
        Set matches = Nothing
        Set tempReg = Nothing
    End Sub

    '类排错模式——在类释放之前应用此方法可以随时查看你截获的内容HTML代码和页面显示效果/方法

    Public Sub debug()
        Dim tempstr
        tempstr = "<SCRIPT>function runEx(){var winEx2 = window.open("""", ""winEx2"", ""width=500,height=300,status=yes,menubar=no,scrollbars=yes,resizable=yes""); winEx2.document.open(""text/html"", ""replace""); winEx2.document.write(unescape(event.srcElement.parentElement.children[0].value)); winEx2.document.close(); }function saveFile(){var win=window.open('','','top=10000,left=10000');win.document.write(document.all.asdf.innerText);win.document.execCommand('SaveAs','','javascript.htm');win.close();}</SCRIPT><center><TEXTAREA id=asdf name=textfield rows=32  wrap=VIRTUAL cols=""120"">"&value_&"</TEXTAREA><BR><BR><INPUT name=Button onclick=runEx() type=button value=""查看效果"">&nbsp;&nbsp;<INPUT name=Button onclick=asdf.select() type=button value=""全选"">&nbsp;&nbsp;<INPUT name=Button onclick=""asdf.value=''"" type=button value=""清空"">&nbsp;&nbsp;<INPUT onclick=saveFile(); type=button value=""保存代码""></center>"
        'response.Write(tempstr)
		document.Write tempstr
    End Sub

End Class


