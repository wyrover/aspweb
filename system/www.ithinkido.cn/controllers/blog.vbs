Class Blog
	Function add()
		
        Import "system/helpers/date_helper.vbs" 		
          
		Dim post_tag, post_taglist, sql, categoryId, categoryId2, tag_id, myPageCreate

		Set myPageCreate = LoadModel("PageCreate")	

		If Request.Form("btnSaveBlog") <> "" Or Request.Form("btnSaveDraft") <> "" Then

			Dim returnURL
			If Request.Cookies("ReturnURL") <> "" Then
				returnURL = Request.Cookies("ReturnURL")
			Else
				returnURL = "index.asp?/blog/add/" & segment(3)
			End If

			Dim IsTop
			IsTop = IIf(Request.Form("chkIsTop") = "on", "1", "0")

				

			If segment(3) <> "" And segment(3) <> "0" Then				
				Set rs = Db.CreateRS()
				sql = "SELECT * FROM blog_blogs WHERE ID =" & segment(3)
				rs.Open sql, Db.ConnectionString, 1, 2
				categoryId = rs("BlogCategoryID")
				categoryId2 = CID(Request.Form("ddlCategories"))
				rs("BlogCategoryID") = CID(Request.Form("ddlCategories"))
				strTitle = CheckStr(Trim(Request.Form("txtTitle")))
				If strLength(strTitle) < 1 Then
					Call MsgBox2("标题字数不能为空", 0, "0")	
				End If
				
				rs("Title") = strTitle
				rs("Author") = CID(Request.Form("ddlAuthor"))				
				rs("IsShow") = CID(Request.Form("ddlAttributes"))
                rs("IsTop") = IsTop
				rs("Alias") = IIf(Request.Form("txtAlias") = "", "0", Request.Form("txtAlias"))   

				tempTags = Split(CheckStr(Request.Form("txtTags"))," ")					 
				Set mytag = LoadModel("tag")
				
				post_taglist = ""
					
				'添加新的Tag				
				
				For Each post_tag In tempTags

					
					If Len(Trim(post_tag))>0 Then
						tag_id = mytag.insert(CheckStr(trim(post_tag)))
						post_taglist = post_taglist & "{" & tag_id & "}"						
						Call myPageCreate.create_tag_archive_by_id(tag_id)						
					End if
				Next
				
				Call mytag.Tags(2)
				Set mytag = Nothing
				
				rs("Tags") = post_taglist

	
				rs("Introduce") = HTMLEncode(Request.Form("txtIntroduce"))
				rs("Content") = HTMLEncode(Request.Form("txtContent"))
				rs("UpdateTime") = Now
				rs.Update
				rs.Close
				Set rs = Nothing

				
				Call myPageCreate.createblogcontentbyid(segment(3))
				If categoryId = categoryId2 Then
					Call myPageCreate.createArchiveByCategoryId(categoryId)
				Else
					Call myPageCreate.createArchiveByCategoryId(categoryId)
					Call myPageCreate.createArchiveByCategoryId(categoryId2)
				End If

				Set myPageCreate = Nothing

				Call MsgBox2("修改成功！",1, returnURL)
			Else
				'Response.Write server.urlencode(Replace(Trim(Request.Form("txtTitle")), " ", "-"))
				'Response.Write URLDecode(URLEncode(Request.Form("txtTitle")))
				'Response.End

			


				Dim id
				categoryId2 = CID(Request.Form("ddlCategories"))
				Set rs = Db.CreateRS()
				sql = "SELECT * FROM blog_blogs WHERE 1 > 2"
				rs.Open sql, Db.ConnectionString, 1, 2
				rs.AddNew
					rs("BlogCategoryID") = CID(Request.Form("ddlCategories"))
					strTitle = CheckStr(Trim(Request.Form("txtTitle")))
					If strLength(strTitle) < 1 Then
						Call MsgBox2("标题字数不能为空", 0, "0")	
					End If

					rs("Title") = strTitle
					rs("Author") = CID(Request.Form("ddlauthor"))
					
					rs("DateNode") = DateToStr(Now, "Y-m") & "-1" 
					rs("DayNode") = DateToStr(Now, "Y-m-d")
					rs("IsShow") = CID(Request.Form("ddlAttributes"))
					rs("IsTop") = IsTop

					
					rs("Alias") = IIf(Request.Form("txtAlias") = "", 0, Request.Form("txtAlias"))   
					
					tempTags = Split(CheckStr(Request.Form("txtTags"))," ")					 
					Set mytag = LoadModel("tag")

					
					post_taglist = ""
					
					'添加新的Tag
					For Each post_tag In tempTags
						If Len(Trim(post_tag))>0 Then
							tag_id = mytag.insert(CheckStr(trim(post_tag)))
							post_taglist = post_taglist & "{" & tag_id & "}"						
							'Call myPageCreate.create_tag_archive_by_id(tag_id)
						End if
					Next
				
					Call mytag.Tags(2)
					Set mytag = Nothing


						


					rs("Tags") = post_taglist
					rs("Introduce") = HTMLEncode(Request.Form("txtIntroduce"))
					rs("Content") = HTMLEncode(Request.Form("txtContent"))
					rs("PostTime") = Now
				rs.Update
					id = rs("ID")	
				rs.Close
				Set rs = Nothing
				

			
				Call myPageCreate.createblogcontentbyid(id)
				'Call myPageCreate.createArchiveByCategoryId(categoryId2)
				Set myPageCreate = Nothing	
				
				Call MsgBox2("添加成功！", 1, returnURL)
			End If			
		End If


		Import "system/libraries/fckeditor.vbs"

	
		Dim sBasePath
		sBasePath = config("base_url") & "system/fckeditor/"
		Dim oFCKeditor
		Set oFCKeditor = New FCKeditor
		oFCKeditor.BasePath	= sBasePath
		oFCKeditor.Config("AutoDetectLanguage") = False
		oFCKeditor.Config("DefaultLanguage")    = "zh-cn"
		oFCKeditor.Config("TabSpaces") = 8		
		oFCKeditor.Height = 450



		Dim oFCKeditor1
        Set oFCKeditor1 = New FCKeditor
		oFCKeditor1.BasePath	= sBasePath
		oFCKeditor1.Height="150"
		oFCKeditor1.ToolbarSet="Basic"
		oFCKeditor1.Config("AutoDetectLanguage") = False
		oFCKeditor1.Config("DefaultLanguage")    = "zh-cn"
		
		
		Dim strTagID, strTag, rs3

		If segment(3) <> "0" Then
			d("action") = "index.asp?/blog/add/" & segment(3)
			d("action2") = "index.asp?/ajax/blog_add/" & segment(3)
			sql = "SELECT * FROM blog_blogs WHERE ID = " & segment(3)
			Set rs = Db.Execute(sql)
			If Not rs.Eof Then
				d("Title") = rs("Title")
                d("Alias") = IIf(rs("Alias") = "0", "", rs("Alias"))                        
				strTagID = Replace(rs("Tags"), "}{", ",")
				strTagID = Mid(strTagID, 2, Len(strTagID) - 2)
				strTag = ""
				Set rs3 = Db.Execute("SELECT * FROM blog_tags WHERE tag_id IN (" & strTagID & ")")
				While Not rs3.Eof
					strTag = strTag & rs3("tag_name") & " "
					rs3.MoveNext
				Wend

				d("Tags") = strTag
				sql2 = "SELECT * FROM blog_categories ORDER BY Sort ASC"
				strBuffer = ""
				Set rs2 = Db.Execute(sql2)
				While Not rs2.Eof
					If rs2("ID") = rs("BlogCategoryID") Then
						strBuffer = strBuffer & "<option value=""" & rs2("ID") &""" selected=""ture"">" & rs2("BlogCategoryName") & "</option>"
					Else
						strBuffer = strBuffer & "<option value=""" & rs2("ID") &""">" & rs2("BlogCategoryName") & "</option>"
					End If
					rs2.MoveNext
				Wend



				d("categories") = BuildCategoryTree2(0, "",  rs("BlogCategoryID"))

				d("attributes") = rs("IsShow")

				sql2 = "SELECT * FROM blog_author ORDER BY ID ASC"
				strBuffer = ""
				Set rs2 = Db.Execute(sql2)
				While Not rs2.Eof
					If rs2("ID") = rs("author") Then
						strBuffer = strBuffer & "<option value=""" & rs2("ID") &""" selected=""ture"">" & rs2("author") & "</option>"
					Else
						strBuffer = strBuffer & "<option value=""" & rs2("ID") &""">" & rs2("author") & "</option>"
					End If
					rs2.MoveNext
				Wend
				d("author") = strBuffer

				

				
				oFCKeditor1.Value	= iHTMLEncode(rs("Introduce"))
				oFCKeditor.Value	= iHTMLEncode(rs("Content"))

				If Len(rs("Introduce")) > 0 Then
					d("ShowIntroduce") = "checked=""checked"""
					d("ShowIntroduce2") = ""
				Else
					d("ShowIntroduce") = ""
					d("ShowIntroduce2") = "style=""display:none"""
				End If

			End If
		Else
			d("action") = "index.asp?/blog/add"
			d("action2") = "index.asp?/blog/add"
			d("attributes") = "1"
			d("Title") = ""
            d("Alias") = ""
			

			sql = "SELECT * FROM blog_author ORDER BY ID ASC"
			strBuffer = ""
			Set rs = Db.Execute(sql)
			While Not rs.Eof
				strBuffer = strBuffer & "<option value=""" & rs("ID") &""">" & rs("author") & "</option>"
				rs.MoveNext
			Wend
			d("author") = strBuffer

			d("Tags") = ""

			
			d("categories") = BuildCategoryTree2(0, "", 0)
		End If

		d("message") = oFCKeditor.Create("txtContent")
		Set oFCKeditor = Nothing


		

		d("introduce") = oFCKeditor1.Create("txtIntroduce")
		Set oFCKeditor1 = Nothing

		t.Load "blog/add.htm", d		
		Set rs = Nothing
	End Function

	Function del()				
		Import "system/helpers/date_helper.vbs" 		

		Dim returnURL, sql
		If Request.Cookies("ReturnURL") <> "" Then
			returnURL = Request.Cookies("ReturnURL")
		Else
			returnURL = "index.asp?/blog/add/" & segment(3)
		End If
		Dim tempIDArray
		tempIDArray = Split(Request.Form("ID"), ",")

		Dim filename, dirpath, filepath2
		For i = 0 To UBound(tempIDArray)
			sql = "DELETE FROM blog_comments WHERE BlogID = " & tempIDArray(i)
			Db.Execute(sql)

			sql = "SELECT * FROM blog_blogs WHERE ID = " & tempIDArray(i)
			Set rs = Db.Execute(sql)
			If Not rs.Eof Then
				
				filename = IIf(rs("Alias") = "0", rs("ID"), rs("Alias"))		
				dirpath = config("archives") & DateToStr(rs("PostTime"), "Y-m")				
				filepath2 = Server.MapPath(".") & dirpath & "\" & Day(rs("PostTime")) & "-" & filename & ".htm"			
				DeleteFiles filepath2			
			End If			
		Next

		sql = "DELETE FROM blog_blogs WHERE ID IN (" & Request.Form("ID") & ")"
		Db.Execute(sql)
		Call MsgBox2("删除成功！", 1, returnURL)
	End Function


	Function list()
		
		Dim pager, currentpage, classid
		Set pager = New Pager

		If segment(4) = "page" And segment(5) <> "" Then				
			currentpage = CInt(segment(5))						
		Else
			currentpage = 1
		End If

		dim searhtext
		searchtext = IIf(segment(3) <> "0", segment(3), "")
		

		Response.Cookies("ReturnURL") = "index.asp?/blog/list/0/page/" & CStr(currentpage)

		dim sql
		sql = "select COUNT(ID) from blog_blogs WHERE Title LIKE '%%" & searchtext & "%%'"
		set rs = db.Execute(sql)
		Call pager.Init("index.asp?/blog/list/" & searchtext & "/page/", currentpage, rs(0), 20, "")		
		set rs = Nothing			
		d("page1") = pager.getHTML()		
		d("pagesize") = pager.PageSize()
		If currentpage = 1 Then
			d("tableB") = "WHERE A.Title LIKE '%%" & searchtext & "%%'"
		Else
			d("tableB") = "WHERE A.Title LIKE '%%" & searchtext & "%%' AND A.ID NOT IN (SELECT TOP " & pager.PageSize() * (currentpage - 1) & " C.ID FROM blog_blogs C WHERE C.Title LIKE '%%" & searchtext & "%%' ORDER BY C.ID DESC)"
		End If	


		d("custom_home_page") = config("custom_home_page")

		t.Load "blog/list.htm", d
	End Function


	Function nopassed_list()		
		Dim pager, currentpage, classid
		Set pager = New Pager

		If segment(3) = "page" And segment(4) <> "" Then				
			currentpage = CInt(segment(4))						
		Else
			currentpage = 1
		End If

		Response.Cookies("ReturnURL") = "index.asp?/manage/blog_list/page/" & CStr(currentpage)

		dim sql
		sql = "select COUNT(*) from blog_blogs" 
		set rs = db.Execute(sql)
		Call pager.Init("index.asp?/manage/blog_list/page/", currentpage, rs(0), 20, "")
		set rs = Nothing			
		d("page1") = pager.getHTML()		
		d("pagesize") = pager.PageSize()
		If currentpage = 1 Then
			d("tableB") = ""
		Else
			d("tableB") = "AND A.ID NOT IN (SELECT TOP " & pager.PageSize() * (currentpage - 1) & " C.ID FROM blog_blogs C ORDER BY C.ID DESC)"	
		End If	

		t.Load "blog/nopassed_list.htm", d
	End Function




	Function BuildCategoryTree2(parentId, prefix, categoryId)

		Dim sql, rs, reval, ddprefix
		sql = "SELECT * FROM blog_categories WHERE ParentID = " & parentId & " ORDER BY Sort ASC"
		Set rs = Db.Execute(sql)
		
		
		While Not rs.eof
			
			sql2 = "SELECT COUNT(*) AS NUM FROM blog_categories WHERE ParentID = " & rs("ID")  
			Set rs2 = Db.Execute(sql2)

			If prefix <> "" Then
				ddprefix = "&nbsp;&nbsp;&nbsp;&nbsp;" & prefix
			Else
				ddprefix = "|—"
			End If

			If categoryId = rs("ID") Then
				reval = reval & "<option value=""" & rs("ID") &""" selected=""ture"">" & ddprefix & rs("BlogCategoryName") & "</option>"
			Else
				reval = reval & "<option value=""" & rs("ID") &""">" &  ddprefix & rs("BlogCategoryName") & "</option>" & vbCrlf		
			End If

			If Not rs2.Eof AND Int(rs2("NUM")) > 0 Then				
				reval = reval & BuildCategoryTree2(rs("ID"), ddprefix, categoryId)				
			End If
			rs2.Close
			Set rs2 = Nothing
			rs.MoveNext 
		Wend		
		

		Set rs = Nothing
		BuildCategoryTree2 = reval
	
	End Function

End Class 