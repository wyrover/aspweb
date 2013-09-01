Function FormatTag(tags)
	Dim regEx, strBuffer, retVal
	Set regEx = New RegExp
	regEx.IgnoreCase = True
	regEx.Global = True
	regEx.MultiLine = True
	regEx.Pattern = "\{(.*?)\}"
	Set matches = regEx.Execute(tags)
	For Each match In matches
		If match.SubMatches(0) <> "" Then				
			strBuffer =  strBuffer & match.SubMatches(0) & ","
		End If
	Next

	sql = "SELECT * FROM blog_tags WHERE tag_id IN (" & strBuffer & ")"
	Set rs = Db.Execute(sql)
	While Not rs.Eof
		retVal = retVal & " <a href=""index.asp?/home/tag/" & rs("tag_name") & """>" & rs("tag_name") & "</a>"
		rs.MoveNext
	Wend
	FormatTag =  Trim(retVal)
End Function

Function FormatTag2(tags)
	Dim regEx, strBuffer, retVal
	Set regEx = New RegExp
	regEx.IgnoreCase = True
	regEx.Global = True
	regEx.MultiLine = True
	regEx.Pattern = "\{(.*?)\}"
	Set matches = regEx.Execute(tags)
	For Each match In matches
		If match.SubMatches(0) <> "" Then				
			strBuffer =  strBuffer & match.SubMatches(0) & ","
		End If
	Next

	sql = "SELECT * FROM blog_tags WHERE tag_id IN (" & strBuffer & ")"
	Set rs = Db.Execute(sql)
	While Not rs.Eof
		retVal = retVal & " <a href=""index.asp?/home/tag/" & rs("tag_name") & """>" & rs("tag_name") & "</a>"
		rs.MoveNext
	Wend
	FormatTag =  Trim(retVal)
End Function



Function FormatContent(content, id)
	FormatContent =  SplitLines(iHTMLEncode(content), 1) 
End Function

Function CheckLogin()
	Dim username, password
	If Request.Cookies("username") <> "" And Request.Cookies("password") <> ""Then				
		username = Request.Cookies("username")
		password = Request.Cookies("password")		
		
		sql = "select * from member where username = '" & username & "'"
		Set rs = Db.Execute(sql)
		If Not rs.Eof Then
				
			If Md5(password, 16) = rs("password") Then										
				Exit Function
			Else
				Response.Cookies("username").Expires = Date - 1
				Response.Cookies("password").Expires = Date - 1
			End If	
		End If			
	End If
	Call MsgBox2("您没有权限访问该页面!", 0, "0")	
End Function

Function FormatDate(DateTime)
	FormatDate = Year(DateTime) & "年" & Month(DateTime) & "月"
End Function

Function FormatDateArchives(DateTime)
	FormatDateArchives = "?/home/archive/" & Year(DateTime) & "-" & Month(DateTime)
End Function

Function manage_check_login()
	Dim username, password
	If Request.Cookies("username") <> "" And Request.Cookies("password") <> ""Then				
		username = Request.Cookies("username")
		password = Request.Cookies("password")		
		
		sql = "select * from member where username = '" & username & "'"
		Set rs = Db.Execute(sql)
		If Not rs.Eof Then				
			If Md5(password, 16) = rs("password") And rs("IsAdmin") = "1" Then										
				Exit Function
			Else
				Response.Cookies("username").Expires = Date - 1
				Response.Cookies("password").Expires = Date - 1
			End If	
		End If			
	End If
	Response.Redirect("index.asp?/manage/login")	
End Function


Function getdressURL(posttime, name, id)
	Dim retVal, strBuffer
	retVal = ""
	retVal = config("archives") & Year(posttime) & "-" & Month(posttime) & "/" & Day(posttime) & "-"
	strBuffer = IIf(Len(Trim(name)) > 2, CStr(name), CStr(id))
	retVal = retVal & strBuffer	
	getdressURL = retVal & ".htm"
End Function