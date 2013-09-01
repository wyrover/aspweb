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
		retVal = retVal & " <a href=""" & Replace(config("archives"), "\", "/") & rs("tag_name") & ".htm"">" & rs("tag_name") & "</a>"
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

Function getContactBlogs(tags, blogId)
	Dim regEx, strBuffer, retVal, sql
	Set regEx = New RegExp
	regEx.IgnoreCase = True
	regEx.Global = True
	regEx.MultiLine = True
	regEx.Pattern = "\{(.*?)\}"
	Set matches = regEx.Execute(tags)
	For Each match In matches
		If match.SubMatches(0) <> "" Then				
			sql = "SELECT Top 5 * FROM blog_blogs WHERE Tags LIKE '%%{" & match.SubMatches(0) & "}%%' AND iD <> " & blogId
			Set rs = Db.Execute(sql)
			While Not rs.Eof
				retVal = retVal & " <a href=""" & getViewURL(rs("PostTime"), rs("Alias"), rs("ID")) & """>" & rs("Title") & "</a><br />"
				rs.MoveNext
			Wend
			
		End If
	Next

	
	
	getContactBlogs =  Trim(retVal)
End Function


Function FormatContent(content)
	FormatContent =  closeHTML(SplitLines(iHTMLEncode(content), 1))
End Function


Function FormatComment(description)
	Dim comment 
	comment = iHTMLEncode(description)

	If StrLength(comment) < 20 Then
		FormatComment =  Left(iHTMLEncode(comment), 20)
	Else
		FormatComment =  Left(iHTMLEncode(comment), 18) & "……"
	End If
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


Function getViewURL(posttime, name, id)
	Dim retVal, strBuffer
	retVal = ""
	retVal = Replace(config("archives"), "\", "/") & Year(posttime) & "-" & Month(posttime) & "/" & Day(posttime) & "-"
	strBuffer = IIf(Len(Trim(name)) > 2, CStr(name), CStr(id))
	retVal = retVal & strBuffer	
	getViewURL = retVal & ".htm"
End Function

Function getCommentURL(posttime, name, blogid, id)
	Dim retVal, strBuffer
	retVal = ""
	retVal = Replace(config("archives"), "\", "/") & Year(posttime) & "-" & Month(posttime) & "/" & Day(posttime) & "-"
	strBuffer = IIf(Len(Trim(name)) > 2, CStr(name), CStr(blogid))
	retVal = retVal & strBuffer	
	getCommentURL = retVal & ".htm#cmt" & id
End Function


Function getTagURL(tag_name)
	Dim retVal
	retVal = ""
	retVal = Replace(config("archives"), "\", "/") & tag_name		
	getTagURL = retVal & ".htm"
End Function

Function getCategoryURL(categoryname)
	Dim retVal
	retVal = ""
	retVal = Replace(config("archives"), "\", "/") & categoryname & "/"		
	getCategoryURL = retVal 
End Function

Function getBlogCountByCategory(categoryid)
	Dim retVal
	retVal = ""
	sql = "select COUNT(ID) as NUM from blog_blogs where BlogCategoryID = " & categoryid
	Set rs = Db.Execute(sql)
	If Not rs.Eof Then
		retVal = rs("NUM")
	End If
	getBlogCountByCategory = retVal
End Function

Function getBlogListByCategory(categoryid)
	Dim sb
	Set sb = New StringBuilder	
	sql = "SELECT A.*, B.BlogCategoryName, B.EName as categoryename, C.Author as authorname, (SELECT COUNT(ID) FROM blog_comments WHERE BlogID = A.ID) AS CommentCount FROM ((blog_blogs A  INNER JOIN blog_categories B ON B.ID = A.BlogCategoryID) INNER JOIN blog_author C ON C.ID = A.Author) WHERE A.IsShow = '1' AND A.BlogCategoryID = " & categoryid & " ORDER BY A.ID DESC" 
	Set rs = Db.Execute(sql)
	sb.Append "<ul>"
	While Not rs.Eof 
		sb.Append "<li><a href=""" & getViewURL(rs("PostTime"), rs("Alias"), rs("ID")) & """ target=""_blank"">" &  rs("Title") & "</a></li>"
		rs.MoveNext
	Wend	
	sb.Append "</ul>"
	getBlogListByCategory = sb.ToString()
	Set sb = Nothing
End Function


Function getArchiveLinkByMonth2(datenode)
	Dim retVal
	retVal = "<a href=""" & Replace(config("archives"), "\", "/")  & DateToStr(datenode, "Y-m") & "/"">" &  DateToStr(datenode, "Y-m") & "</a>"
	getArchiveLinkByMonth2 = retVal
End Function

Function getArchiveLinkByMonth(datenode)
	Dim retVal
	retVal = "<a href=""" & Replace(config("archives"), "\", "/")  & DateToStr(datenode, "Y-m") & "/"">" &  Date2ChineseRSS(datenode) & "</a>"
	getArchiveLinkByMonth = retVal
End Function

Function getRssURL(categoryname)
	Dim retVal
	retVal = ""
	retVal = Replace(config("archives"), "\", "/") & categoryname	& ".xml"
	getRssURL = retVal 
End Function