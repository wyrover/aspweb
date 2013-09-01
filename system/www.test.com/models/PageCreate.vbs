Class PageCreate

	  Function createblogcontentbyid(id)
		Import "system/helpers/date_helper.vbs"	
		dim reval, sql, filename, dirpath, filepath2

			
		sql = "SELECT * FROM blog_blogs WHERE ID = " & id
		Set rs = Db.Execute(sql)
		If Not rs.Eof Then
			filename = IIf(rs("Alias") = "0", rs("ID"), rs("Alias"))		
			dirpath = config("archives") & DateToStr(rs("PostTime"), "Y-m")
			Call CreateFolder(dirpath)
			filepath2 = Server.MapPath(".") & dirpath & "\" & Day(rs("PostTime")) & "-" & filename & ".htm"			
			sb.Clear
			d("blogId") = rs("ID")
			Call BuildFile(filepath2, p.Parser(t.Load("skins/" & config("skin") & "/view_c.htm", d)), 0)
			rs.MoveNext
		End If		
		sb.Clear	

	End Function



	Function createArchiveByCategoryId(categoryId)
		Import "system/helpers/date_helper.vbs"			

		Dim pager, currentpage, classid, recordcount
		Set pager = New Pager

		
		dim sql, ename
		sql = "select COUNT(*) from blog_blogs WHERE IsShow = '1' AND BlogCategoryID = " & categoryId 
		set rs = db.Execute(sql)
		recordcount = rs(0)	
		rs.Close
		Set rs = Nothing

		sql = "SELECT EName FROM blog_categories WHERE ID = " & categoryId
		set rs = db.Execute(sql)		
		ename = rs(0)
		rs.Close
		Set rs = Nothing

		Dim pagecount
		If recordcount mod 10 = 0 Then
			pagecount = recordcount / 10
		Else
			pagecount = (Int(recordcount / 10)) + 1
		End If		

		filepath = Server.MapPath(".") & config("archives") & ename 	
		Call CreateFolder(config("archives") & ename & "\")
	
		For currentpage = 1 To pagecount
		

		

		
			Call pager.Init("page-", currentpage, recordcount, 10, ".htm")		
			d("page1") = pager.getHTML()		
			d("pagesize") = pager.PageSize()
			If currentpage = 1 Then
				d("tableB") = "WHERE A.IsShow = '1' AND A.BlogCategoryID = " & categoryId
			Else
				d("tableB") = "WHERE A.IsShow = '1' AND A.BlogCategoryID = " & categoryId & " AND A.ID NOT IN (SELECT TOP " & pager.PageSize() * (currentpage - 1) & " C.ID FROM blog_blogs C WHERE BlogCategoryID = " & categoryId & " ORDER BY C.ID DESC)"	
			End If	

		
			filepath2 = filepath & "\page-" & currentpage & ".htm"			
			sb.Clear
			Call BuildFile(filepath2, p.Parser(t.Load("skins/" & config("skin") & "/home_c.htm", d)), 0)		
			
			If currentpage = 1 Then
				filepath2 = Server.MapPath(".") & config("archives") & ename & "\index.htm" 
				sb.Clear
				Call pager.Init(Replace(config("archives"), "\", "/") & ename & "/page-", currentpage, recordcount, 10, ".htm")		
				d("page1") = pager.getHTML()		
				d("pagesize") = pager.PageSize()
				Call BuildFile(filepath2, p.Parser(t.Load("skins/" & config("skin") & "/home_c.htm", d)), 0)			
			End If
		Next
		
		sb.Clear
	End Function


	Function create_tag_archive_by_id(tag_id)
		dim reval, tagname, filepath2
		Import "system/helpers/date_helper.vbs"		


		Set rs = Db.Execute("SELECT * FROM blog_tags WHERE tag_id = " & tag_id)
		If Not rs.Eof Then
			
			filepath2 = Server.MapPath(".") & config("archives") &  rs("tag_name") & ".htm"			
			sb.Clear
			d("tag_id") = "'%%{" & rs("tag_id") & "}%%'"
			d("tag_name") = rs("tag_name")
			Call BuildFile(filepath2, p.Parser(t.Load("skins/" & config("skin") & "/tag_c.htm", d)), 0)
			rs.MoveNext
		End If
		
		sb.Clear			

	End Function


	Function BuildTree(parentId)

		Dim sql, rs, reval
		sql = "SELECT * FROM blog_categories WHERE ParentID = " & parentId & " ORDER BY ID ASC"
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
 
End Class