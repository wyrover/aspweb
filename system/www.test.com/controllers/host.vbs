Class Host
	Function index()
		
		dim sql
		sql = "select COUNT(*) from blog_blogs" 
		set rs = db.Execute(sql)
		recordcount = rs(0)
		rs.Close
		Set rs = Nothing

		If recordcount = 0 Then
			Response.Write "记录数为空，不能生成"
			Response.End
		End If

		Dim pagecount
		If recordcount mod 10 = 0 Then
			pagecount = recordcount / 10
		Else
			pagecount = (Int(recordcount / 10)) + 1
		End If	
		
		d("pagecount") = pagecount


		sql = "SELECT MIN(ID) FROM blog_blogs"
		Set rs = Db.Execute(sql)
		If Not rs.Eof Then		
			d("BeginBlogID") = rs(0)
		End If

		sql = "SELECT MAX(ID) FROM blog_blogs"
		Set rs = Db.Execute(sql)
		If Not rs.Eof Then		
			d("EndBlogID") = rs(0)
		End If

		sql = "SELECT MIN(tag_id) FROM blog_tags"
		Set rs = Db.Execute(sql)
		If Not rs.Eof Then
			d("BeginTagID") = rs(0)
		End If

		sql = "SELECT MAX(tag_id) FROM blog_tags"
		Set rs = Db.Execute(sql)
		If Not rs.Eof Then
			d("EndTagID") = rs(0)
		End If

		sql = "SELECT MIN(ID) FROM blog_categories"
		Set rs = Db.Execute(sql)
		If Not rs.Eof Then	
			d("BeginCategoryID") = rs(0)
		End If

		sql = "SELECT MAX(ID) FROM blog_categories"
		Set rs = Db.Execute(sql)
		If Not rs.Eof Then	
			d("EndCategoryID") = rs(0)
		End If

		d("custom_home_page") = config("custom_home_page")

		t.Load "manage/host_index.htm", d
	End Function

	' 生成内容页面
	Function createblogcontent()
		Import "system/helpers/date_helper.vbs"
		Response.Buffer = true
		Response.ExpiresAbsolute = Now - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"
		dim reval, sql, filename, dirpath
		reval = reval & "<div class=""CNLTreeMenu"" id=""CNLTreeMenu2"">" & vbCrlf	
		reval = reval & BuildTree(0)
		reval = reval & "</div>" & vbCrlf
		d("tree") = reval
		
		
		
		sql = "SELECT * FROM blog_blogs WHERE ID >= " & (Int(segment(3)) - 1) * 100 & " AND ID < " & Int(segment(3)) * 100
		Set rs = Db.Execute(sql)
		While Not rs.Eof 
			filename = IIf(rs("Alias") = "0", rs("ID"), rs("Alias"))		
			dirpath = config("archives") & DateToStr(rs("PostTime"), "Y-m")
			Call CreateFolder(dirpath)
			filepath2 = Server.MapPath(".") & dirpath & "\" & Day(rs("PostTime")) & "-" & filename & ".htm"			
			sb.Clear
			d("blogId") = rs("ID")
			Call BuildFile(filepath2, p.Parser(t.Load("skins/" & config("skin") & "/view_c.htm", d)), 0)
			rs.MoveNext
		Wend
		
		sb.Clear
		sb.Append(filepath2)		
	End Function


	Function BuildTree(parentId)

		Dim sql, rs, reval
		sql = "SELECT * FROM blog_categories WHERE ParentID = " & parentId & " ORDER BY Sort DESC"
		Set rs = Db.Execute(sql)
		reval = reval & "<ul>" & vbCrlf
		While Not rs.eof
			
			sql2 = "SELECT COUNT(*) AS NUM FROM blog_categories WHERE ParentID = " & rs("ID")  
			Set rs2 = Db.Execute(sql2)
			If Not rs2.Eof AND Int(rs2("NUM")) > 0 Then
				reval = reval & "<li><a href=""" & getCategoryURL(rs("EName")) & """>" & rs("BlogCategoryName") & "</a></li>" & vbCrlf
				reval = reval & BuildTree(rs("ID"))			
			Else
				reval = reval & "<li class=""Child""><a href=""" & getCategoryURL(rs("EName")) & """>" & rs("BlogCategoryName") & "</a>  <a href=""" &  getRssURL(rs("EName")) &"""><img src=""/images/rss_icon2.png"" /></a></li>" & vbCrlf
			End If
			rs2.Close
			Set rs2 = Nothing
			rs.MoveNext 
		Wend		
		reval = reval & "</ul>" & vbCrlf
		Set rs = Nothing
		BuildTree = reval
	
	End Function	


	' 生成首页和列表页
	Function createpage2()
		Import "system/helpers/date_helper.vbs"
		Response.Buffer = true
		Response.ExpiresAbsolute = Now - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"	
			
		

		Dim pager, currentpage, classid, recordcount
		Set pager = New Pager

		

		dim sql
		sql = "select COUNT(*) from blog_blogs WHERE IsShow = '1'" 
		set rs = db.Execute(sql)
		recordcount = rs(0)
		rs.Close
		Set rs = Nothing

		Dim pagecount
		If recordcount mod 10 = 0 Then
			pagecount = recordcount / 10
		Else
			pagecount = (Int(recordcount / 10)) + 1
		End If		

		filepath = Server.MapPath(".") & config("archives") 		
		Call CreateFolder(config("archives"))
	
		currentpage = Int(segment(3))
	
		Call pager.Init("page", currentpage, recordcount, 10, ".htm")		
		d("page1") = pager.getHTML()		
		d("pagesize") = pager.PageSize()
		If currentpage = 1 Then
			d("tableB") = "WHERE A.IsShow = '1' "
		Else
			d("tableB") = "WHERE A.IsShow = '1' AND A.ID NOT IN (SELECT TOP " & pager.PageSize() * (currentpage - 1) & " C.ID FROM blog_blogs C WHERE C.IsShow = '1' ORDER BY C.ID DESC)"	
		End If	

	
		filepath2 = filepath & "page" & currentpage & ".htm"			
		sb.Clear
		Call BuildFile(filepath2, p.Parser(t.Load("skins/" & config("skin") & "/home_c.htm", d)), 0)		
		
		If currentpage = 1 Then
			filepath2 = Server.MapPath(".") & "\" & Application_PATH & ".htm" 
			sb.Clear
			Call pager.Init(Replace(config("archives"), "\", "/") & "page", currentpage, recordcount, 10, ".htm")		
			d("page1") = pager.getHTML()		
			d("pagesize") = pager.PageSize()
			Call BuildFile(filepath2, p.Parser(t.Load("skins/" & config("skin") & "/home_c.htm", d)), 0)			
		End If
		
		sb.Clear
		sb.Append(filepath2)	
	End Function


	' 生成tag索引页
	Function create_tag_archive()
		dim reval, tagname

		Import "system/helpers/date_helper.vbs"
		Response.Buffer = true
		Response.ExpiresAbsolute = Now - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"	

		reval = reval & "<div class=""CNLTreeMenu"" id=""CNLTreeMenu2"">" & vbCrlf	
		reval = reval & BuildTree(0)
		reval = reval & "</div>" & vbCrlf
		d("tree") = reval
		
	
		Set rs = Db.Execute("SELECT * FROM blog_tags WHERE tag_id >= " & (Int(segment(3)) - 1) * 100 & " AND tag_id < " & Int(segment(3)) * 100)
		While Not rs.Eof 
			
			filepath2 = Server.MapPath(".") & config("archives") &  rs("tag_name") & ".htm"			
			sb.Clear
			d("tag_id") = "'%%{" & rs("tag_id") & "}%%'"
			d("tag_name") = rs("tag_name")
			Call BuildFile(filepath2, p.Parser(t.Load("skins/" & config("skin") & "/tag_c.htm", d)), 0)
			rs.MoveNext
		Wend
		
		sb.Clear
		sb.Append(filepath2)			

	End Function

	Function createArchiveByMonth()
		dim reval, tagname

		Import "system/helpers/date_helper.vbs"
		Response.Buffer = true
		Response.ExpiresAbsolute = Now - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"
		
	
		Set rs = Db.Execute("SELECT DISTINCT(DateNode) FROM blog_blogs")
		While Not rs.Eof 			
			filepath2 = Server.MapPath(".") & config("archives") &  DateToStr(rs("DateNode"), "Y-m") & "\index.htm"			
			sb.Clear
			d("datenode") = Date2ChineseRSS(rs("DateNode"))
			d("tableB") = "WHERE A.datenode = #" & rs("DateNode") & "#"			
			Call BuildFile(filepath2, p.Parser(t.Load("skins/" & config("skin") & "/archives_c.htm", d)), 0)
			rs.MoveNext
		Wend
		
		sb.Clear
		sb.Append(filepath2)	
	End Function

	Function createArchiveByCategory()
		Import "system/helpers/date_helper.vbs"
		Response.Buffer = true
		Response.ExpiresAbsolute = Now - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"	
			
		dim reval

		reval = GetCacheValue("dtree")
		If reval = "" Then 	
			reval = reval & "<div class=""CNLTreeMenu"" id=""CNLTreeMenu2"">" & vbCrlf	
			reval = reval & BuildTree(0)
			reval = reval & "</div>" & vbCrlf
			SetCacheValue "dtree", reval, 180
		End If

		d("tree") = reval

		Dim pager, currentpage, classid, recordcount
		Set pager = New Pager

		
		dim sql
		sql = "select COUNT(*) from blog_blogs WHERE IsShow = '1' AND BlogCategoryID = " & Request.Form("id")
		set rs = db.Execute(sql)
		recordcount = rs(0)
		rs.Close
		Set rs = Nothing

		Dim pagecount
		If recordcount mod 10 = 0 Then
			pagecount = recordcount / 10
		Else
			pagecount = (Int(recordcount / 10)) + 1
		End If		

		filepath = Server.MapPath(".") & config("archives") & Request.Form("ename") 	
		Call CreateFolder(config("archives") & Request.Form("ename") & "\")
	
		For currentpage = 1 To pagecount
		

		

		
			Call pager.Init("", currentpage, recordcount, 10, ".htm")		
			d("page1") = pager.getHTML()		
			d("pagesize") = pager.PageSize()
			If currentpage = 1 Then
				d("tableB") = "WHERE A.IsShow = '1' AND A.BlogCategoryID = " & Request.Form("id")
			Else
				d("tableB") = "WHERE A.IsShow = '1' AND A.BlogCategoryID = " & Request.Form("id") & " AND A.ID NOT IN (SELECT TOP " & pager.PageSize() * (currentpage - 1) & " C.ID FROM blog_blogs C WHERE BlogCategoryID = " & Request.Form("id") & " ORDER BY C.ID DESC)"	
			End If	

		
			filepath2 = filepath & "\" & currentpage & ".htm"			
			sb.Clear
			Call BuildFile(filepath2, p.Parser(t.Load("skins/" & config("skin") & "/home_c.htm", d)), 0)		
			
			If currentpage = 1 Then
				filepath2 = Server.MapPath(".") & config("archives") & Request.Form("ename") & "\index.htm" 
				sb.Clear
				Call pager.Init(Replace(config("archives"), "\", "/") & Request.Form("ename") & "/", currentpage, recordcount, 10, ".htm")		
				d("page1") = pager.getHTML()		
				d("pagesize") = pager.PageSize()
				Call BuildFile(filepath2, p.Parser(t.Load("skins/" & config("skin") & "/home_c.htm", d)), 0)			
			End If
		Next
		
		sb.Clear
		sb.Append(filepath2)	


	End Function

	Function createFeed()
		Import "system/helpers/date_helper.vbs"
		Response.Buffer = true
		Response.ExpiresAbsolute = Now - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"	

		Dim sCrLf, fileName, sql, sb2
		Set sb2 = New StringBuilder
		categoryId = segment(3)				
		sCrLf = chr(13) & chr(10) 		

		sb2.Append "<?xml version=""1.0"" encoding=""" & config("rss_encoding") & """?>" & vbCrLf
		sb2.Append "<rss version=""2.0"">" & sCrLf
		sb2.Append "<channel>" & sCrLf
		sb2.Append "<title><![CDATA[" & config("site_name") & "]]></title>" & sCrLf
		sb2.Append "<link>" & config("base_url") & "</link>" & sCrLf	
		sb2.Append "<description><![CDATA[" & config("site_description") & "]]></description>" & sCrLf		
		sb2.Append "<language>" & config("rss_language") & "</language>" & sCrLf
		sb2.Append "<generator>" & config("site_name") & "</generator>" & sCrLf
		sb2.Append "<copyright><![CDATA[Copyright " & Year(Now()) & ", " & config("site_name") & "]]></copyright>" & sCrLf		
		sb2.Append "<lastBuildDate>" & DateToStr2(Now(),"w,d m y H:I:S") & "</lastBuildDate>" & sCrLf
		sb2.Append "<pubDate>" & DateToStr2(Now(),"w,d m y H:I:S") & "</pubDate>" & sCrLf
		sb2.Append "<ttl>60</ttl>" & sCrLf

		
	
		If segment(3) = "0" Then
			sql = "SELECT TOP 50 A.*, B.BlogCategoryName, B.Ename as categoryename, C.Author as authorname FROM ((blog_blogs A  INNER JOIN blog_categories B ON B.ID = A.BlogCategoryID) INNER JOIN blog_author C ON C.ID = A.Author) ORDER BY A.ID DESC"
			fileName = Server.MapPath(".") & config("archives") & "feed.xml"	
		Else
			sql = "SELECT TOP 50 A.*, B.BlogCategoryName, B.Ename as categoryename, C.Author as authorname FROM ((blog_blogs A  INNER JOIN blog_categories B ON B.ID = A.BlogCategoryID) INNER JOIN blog_author C ON C.ID = A.Author) WHERE BlogCategoryID = " & segment(3) & " ORDER BY A.ID DESC"
			
		End If

		Set rs = Db.Execute(sql)	

		While Not rs.Eof

			If segment(3) <> "0" Then
				fileName = Server.MapPath(".") & config("archives") & rs("categoryename") & ".xml"	
			End If

			sb2.Append "<item>" & sCrLf
			sb2.Append "<title><![CDATA[" & iHTMLEncode(rs("title")) & "]]></title>" & sCrLf
			sb2.Append "<link>" & Left(config("base_url"), Len(config("base_url")) - 1) & getViewURL(PostTime, Alias, ID) & "</link>" & sCrLf
			sb2.Append "<category><![CDATA[" & rs("BlogCategoryName") & "]]></category>" & sCrLf
			sb2.Append "<author>" & rs("authorname") & "</author>" & sCrLf
			sb2.Append "<pubDate>" & DateToStr2(rs("PostTime"),"w,d m y H:I:S") & " </pubDate>" & sCrLf
			sb2.Append "<description><![CDATA[" & iHTMLEncode(rs("Content")) & "]]></description>" & sCrLf
			sb2.Append "</item>" & sCrLf 

			rs.MoveNext
		Wend

		sb2.Append "</channel></rss>"

		Call BuildFile(fileName, sb2.ToString(), 0)				
		Set sb2 = Nothing
		
		sb.Append fileName

	End Function

	Function newsite()
		Dim site
		site = IIf(Request.Form("site") <> "", Request.Form("site"), "")
		If site = "" Then
			Echo "请填写站点域名"
			Response.End
		Else
			Set fso = CreateObject("Scripting.FileSystemObject")
			Call CreateFolder("system\" & CStr(site))
			fso.CopyFolder Server.MapPath(".") & "\system\www.test.com\*", Server.MapPath(".") & "\system\" & CStr(site) & "\"
			Echo site & "创建成功"
		End If
		
	End Function

	Function createcustomhomepage()
		Import "system/helpers/date_helper.vbs"	
		Response.Buffer = true
		Response.ExpiresAbsolute = Now - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"	
		dim reval, sql, filename, dirpath, filepath2

			
		sql = "SELECT * FROM blog_blogs WHERE ID = " & segment(3)
		Set rs = Db.Execute(sql)
		If Not rs.Eof Then			
			filepath2 = Server.MapPath(".") & "\" & Application_PATH & ".htm" 		
			sb.Clear
			d("blogId") = rs("ID")
			Call BuildFile(filepath2, p.Parser(t.Load("skins/" & config("skin") & "/custom_home_c.htm", d)), 0)
			rs.MoveNext
		End If		
		sb.Clear	



			Dim retVal		
			filepath = Server.MapPath(".") & "\system\" & Application_PATH & "\config\config.vbs"			
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

			Dim regEx
			Set regEx = New RegExp
			regEx.IgnoreCase = True
			regEx.Global = True
			regEx.MultiLine = True

			regEx.Pattern = "^config\(\""custom_home_page\""\).*\=.*?$"
			retVal = regEx.Replace(retVal, "config(""custom_home_page"") = """ & segment(3) & """")		
			
			Set regEx = Nothing				
			Call BuildFile(filepath, retVal, 0)
		


	End Function

End Class
