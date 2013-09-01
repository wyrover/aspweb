Import "system/helpers/date_helper.vbs"



Class Home

' ****************************************************
	Function index()
		'If StartURLCache() = FALSE Then

			'Response.Write GetUrl2
			'Response.End 

			
			
			dim reval
			reval = reval & "<div class=""CNLTreeMenu"" id=""CNLTreeMenu2"">" & vbCrlf	
			reval = reval & BuildTree(0)
			reval = reval & "</div>" & vbCrlf
			d("tree") = reval

			Dim pager, currentpage, classid
			Set pager = New Pager

			If segment(3) = "page" And segment(4) <> "" Then				
				currentpage = CInt(segment(4))						
			Else
				currentpage = 1
			End If

			dim sql
			sql = "select COUNT(*) from blog_blogs" 
			set rs = db.Execute(sql)
			Call pager.Init("index.asp?/home/index/page/", currentpage, rs(0), 10, "")
			set rs = Nothing			
			d("page1") = pager.getHTML()		
			d("pagesize") = pager.PageSize()
			If currentpage = 1 Then
				d("tableB") = ""
			Else
				d("tableB") = "WHERE A.ID NOT IN (SELECT TOP " & pager.PageSize() * (currentpage - 1) & " C.ID FROM blog_blogs C ORDER BY C.ID DESC)"	
			End If	


			t.Load "home.htm", d			
		'End If
		
	End Function

' ******************************************************************
	Function category()
		
			
			dim reval
			reval = reval & "<div class=""CNLTreeMenu"" id=""CNLTreeMenu2"">" & vbCrlf	
			reval = reval & BuildTree(0)
			reval = reval & "</div>" & vbCrlf
			d("tree") = reval

			Dim pager, currentpage, classid
			Set pager = New Pager

			If segment(4) = "page" And segment(5) <> "" Then				
				currentpage = CInt(segment(5))						
			Else
				currentpage = 1
			End If			

			dim sql
			sql = "select COUNT(A.ID) from blog_blogs A WHERE A.BlogCategoryID = " & segment(3)
			set rs = db.Execute(sql)
			Call pager.Init("index.asp?/home/category/" & segment(3) & "/page/", currentpage, rs(0), 10, "")
			set rs = Nothing
			

			sql = "SELECT BlogCategoryName FROM blog_categories WHERE ID = " & segment(3)
			Set rs = Db.Execute(sql)
			d("categoryname") = rs(0)
			Set rs = Nothing


			d("categoryId") = segment(3)
			d("page1") = pager.getHTML()		
			d("pagesize") = pager.PageSize()
			If currentpage = 1 Then
				d.Add "tableB", ""
			Else
				d.Add "tableB", " AND A.ID NOT IN (SELECT TOP " & pager.PageSize() * (currentpage - 1) & " C.ID FROM blog_blogs C ORDER BY C.ID DESC)"	
			End If	


			t.Load "category.htm", d		
	End Function

' ******************************************************************
	Function archives()
		
			
			dim reval
			reval = reval & "<div class=""CNLTreeMenu"" id=""CNLTreeMenu2"">" & vbCrlf	
			reval = reval & BuildTree(0)
			reval = reval & "</div>" & vbCrlf
			d("tree") = reval

			If segment(3) <> "0" Then
				d("datetime") = segment(3)		
				t.Load "update.htm", d	
			Else
								
				t.Load "archives.htm", d
			End If
	End Function
' ******************************************************************
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
' ******************************************************************
	Function login()		
		If Request.Form("txtUsername") <> "" Then
			Set rs = Db.Execute("SELECT * FROM template_list")
			If Not rs.eof Then
				Response.Write rs("title")				
			Else
				Response.Write "ddddddd"
			End If

			Exit Function
		End If

		d.Add "result", "<cms:list><sql>SELECT * FROM template_list</sql><template><li>$title$</li></template></cms:list>"
		t.Load "login.htm", d
	End Function
' ******************************************************************
	Function logout()
		
		Redirect("home/index")
	End Function
' ******************************************************************
	Function download()

		dim reval
		reval = reval & "<div class=""CNLTreeMenu"" id=""CNLTreeMenu2"">" & vbCrlf	
		reval = reval & BuildTree(0)
		reval = reval & "</div>" & vbCrlf
		d("tree") = reval

		Dim pager, currentpage, classid
		Set pager = New Pager

		If segment(3) = "page" And segment(4) <> "" Then				
			currentpage = CInt(segment(4))						
		Else
			currentpage = 1
		End If			

		dim sql
		sql = "select COUNT(A.ID) from blog_blogs A WHERE A.BlogCategoryID = 17" 
		set rs = db.Execute(sql)
		Call pager.Init("index.asp?/home/download/page/", currentpage, rs(0), 10, "")
		set rs = Nothing
		

		sql = "SELECT BlogCategoryName FROM blog_categories WHERE ID = 17" 
		Set rs = Db.Execute(sql)
		d("categoryname") = rs(0)
		Set rs = Nothing


		d("categoryId") = 17
		d("page1") = pager.getHTML()		
		d("pagesize") = pager.PageSize()
		If currentpage = 1 Then
			d.Add "tableB", ""
		Else
			d.Add "tableB", " AND A.ID NOT IN (SELECT TOP " & pager.PageSize() * (currentpage - 1) & " C.ID FROM blog_blogs C ORDER BY C.ID DESC)"	
		End If	


		t.Load "download.htm", d
	End Function
' ******************************************************************
	Function search()				
		dim reval
		reval = reval & "<div class=""CNLTreeMenu"" id=""CNLTreeMenu2"">" & vbCrlf	
		reval = reval & BuildTree(0)
		reval = reval & "</div>" & vbCrlf
		d("tree") = reval

		Dim pager, currentpage, classid
		Set pager = New Pager

		If segment(4) = "page" And segment(5) <> "" Then				
			currentpage = CInt(segment(5))						
		Else
			currentpage = 1
		End If

		dim searhtext
		searchtext = IIf(segment(3) <> "0", segment(3), "")

		dim sql
		sql = "select COUNT(ID) from blog_blogs WHERE Title LIKE '%%" & searchtext & "%%'"
		set rs = db.Execute(sql)
		Call pager.Init("index.asp?/home/search/" & searchtext & "/page/", currentpage, rs(0), 10, "")
		set rs = Nothing			
		d("page1") = pager.getHTML()		
		d("pagesize") = pager.PageSize()
		If currentpage = 1 Then
			d("tableB") = "WHERE A.Title LIKE '%%" & searchtext & "%%'"
		Else
			d("tableB") = "WHERE A.Title LIKE '%%" & searchtext & "%%' AND A.ID NOT IN (SELECT TOP " & pager.PageSize() * (currentpage - 1) & " C.ID FROM blog_blogs C WHERE C.Title LIKE '%%" & searchtext & "%%' ORDER BY C.ID DESC)"	
		End If	


		t.Load "home.htm", d			
	End Function

	
' ******************************************************************
	Function tag()
		dim reval
		reval = reval & "<div class=""CNLTreeMenu"" id=""CNLTreeMenu2"">" & vbCrlf	
		reval = reval & BuildTree(0)
		reval = reval & "</div>" & vbCrlf
		d("tree") = reval

		Dim pager, currentpage, classid
		Set pager = New Pager

		If segment(4) = "page" And segment(5) <> "" Then				
			currentpage = CInt(segment(5))						
		Else
			currentpage = 1
		End If

		dim sql, tagid
		sql = "SELECT tag_id FROM blog_tags WHERE tag_name = '" & segment(3) & "'"
		Set rs = Db.Execute(sql)
		If Not rs.Eof Then
			tagid = rs("tag_id")
		End If
		rs.Close
		Set rs = Nothing



		sql = "select COUNT(ID) from blog_blogs WHERE Tags LIKE '%{" & tagid & "}%'" 
		set rs = db.Execute(sql)
		Call pager.Init("index.asp?/home/tag/" & segment(3) & "/page/", currentpage, rs(0), 5, "")
		set rs = Nothing			
		d("page1") = pager.getHTML()		
		d("pagesize") = pager.PageSize()
		If currentpage = 1 Then
			d.Add "tableB", "WHERE A.Tags LIKE '%{" & tagid & "}%'"
		Else
			d.Add "tableB", "WHERE A.ID NOT IN (SELECT TOP " & pager.PageSize() * (currentpage - 1) & " C.ID FROM blog_blogs C WHERE C.Tags LIKE '%{" & tagid & "}%' ORDER BY C.ID DESC)"	
		End If	


		t.Load "tag.htm", d		
	End Function
' ******************************************************************
	Function tags()
		dim reval
		reval = reval & "<div class=""CNLTreeMenu"" id=""CNLTreeMenu2"">" & vbCrlf	
		reval = reval & BuildTree(0)
		reval = reval & "</div>" & vbCrlf
		d("tree") = reval

		t.Load "tags.htm", d
	End Function
' ******************************************************************
	Function submit()
		Call manage_check_login
		Dim post_tag, post_taglist
		If Request.Form("btnSaveBlog") <> "" Then

			If segment(3) <> "" And segment(3) <> "0" Then
				Set rs = Db.CreateRS()
				sql = "SELECT * FROM blog_blogs WHERE ID =" & segment(3)
				rs.Open sql, Db.ConnectionString, 1, 2
			
				rs("BlogCategoryID") = CID(Request.Form("ddlCategories"))
				strTitle = HTMLEncode(Trim(Request.Form("txtTitle")))
				If strLength(strTitle) < 6 Then
					Call MsgBox2("标题字数不能少于6个", 0, "0")	
				End If

				rs("Title") = strTitle
				rs("Author") = CID(Request.Form("ddlAuthor"))
				rs("Source") = CID(Request.Form("ddlfrom"))
				rs("IsShow") = CID(Request.Form("ddlAttributes"))

				tempTags = Split(CheckStr(Request.Form("txtTags"))," ")					 
				Set mytag = LoadModel("tag")
				
				post_taglist = ""
					
				'添加新的Tag
				For Each post_tag In tempTags
					If Len(Trim(post_tag))>0 Then
						post_taglist = post_taglist & "{" & mytag.insert(CheckStr(trim(post_tag))) & "}"
					End if
				Next
			
				Call mytag.Tags(2)
				Set mytag = Nothing

				rs("Tags") = post_taglist

	
				rs("Content") = HTMLEncode(Request.Form("txtContent"))
				rs("PostTime") = Now
				rs.Update
				rs.Close

				Dim returnURL
				If Request.Cookies("ReturnURL") <> "" Then
					returnURL = Request.Cookies("ReturnURL")
				Else
					returnURL = "index.asp?/home/submit/" & segment(3)
				End If

				

				Call MsgBox2("修改成功！",1, returnURL)
			Else
				Set rs = Db.CreateRS()
				sql = "SELECT * FROM blog_blogs WHERE 1 > 2"
				rs.Open sql, Db.ConnectionString, 1, 2
				rs.AddNew
					rs("BlogCategoryID") = CID(Request.Form("ddlCategories"))
					strTitle = HTMLEncode(Trim(Request.Form("txtTitle")))
					If strLength(strTitle) < 6 Then
						Call MsgBox2("标题字数不能少于6个", 0, "0")	
					End If

					rs("Title") = strTitle
					rs("Author") = Request.Cookies("username")
					rs("IsShow") = CID(Request.Form("ddlAttributes"))

					
					tempTags = Split(CheckStr(Request.Form("txtTags"))," ")					 
					Set mytag = LoadModel("tag")

					
					post_taglist = ""
						
					'添加新的Tag
					For Each post_tag In tempTags
						If Len(Trim(post_tag))>0 Then
							post_taglist = post_taglist & "{" & mytag.insert(CheckStr(trim(post_tag))) & "}"
						End if
					Next
				
					Call mytag.Tags(2)
					Set mytag = Nothing

					rs("Tags") = post_taglist
					rs("Content") = HTMLEncode(Request.Form("txtContent"))
					rs("PostTime") = Now
				rs.Update
				rs.Close
				Call MsgBox2("添加成功！",1,"index.asp?/home/submit")
			End If			
		End If


		Import "system/libraries/fckeditor.vbs"

		dim reval
		reval = reval & "<div class=""CNLTreeMenu"" id=""CNLTreeMenu2"">" & vbCrlf	
		reval = reval & BuildTree(0)
		reval = reval & "</div>" & vbCrlf
		d("tree") = reval

		Dim sBasePath
		sBasePath = config("base_url") & "fckeditor/"
		Dim oFCKeditor
		Set oFCKeditor = New FCKeditor
		oFCKeditor.BasePath	= sBasePath
		oFCKeditor.Config("AutoDetectLanguage") = False
		oFCKeditor.Config("DefaultLanguage")    = "zh-cn"
		oFCKeditor.Config("TabSpaces") = 8		
		oFCKeditor.Height = 450
		

		If segment(3) <> 0 Then
			d("action") = "index.asp?/home/submit/" & segment(3)
			sql = "SELECT * FROM blog_blogs WHERE ID = " & segment(3)
			Set rs = Db.Execute(sql)
			If Not rs.Eof Then
				d("Title") = rs("Title")
				d("Tags") = rs("Tags")
				sql2 = "SELECT * FROM blog_categories ORDER BY ID ASC"
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
				d("categories") = strBuffer

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

				sql2 = "SELECT * FROM blog_from ORDER BY ID ASC"
				strBuffer = ""
				Set rs2 = Db.Execute(sql2)
				While Not rs2.Eof
					If rs2("ID") = rs("source") Then
						strBuffer = strBuffer & "<option value=""" & rs2("ID") &""" selected=""ture"">" & rs2("source") & "</option>"
					Else
						strBuffer = strBuffer & "<option value=""" & rs2("ID") &""">" & rs2("source") & "</option>"
					End If
					rs2.MoveNext
				Wend
				d("from") = strBuffer

				

				oFCKeditor.Value	= iHTMLEncode(rs("Content"))
			End If
		Else
			d("action") = "index.asp?/home/submit"
			d("Title") = ""
			sql = "SELECT * FROM blog_from ORDER BY ID ASC"
			Set rs = Db.Execute(sql)
			While Not rs.Eof
				strBuffer = strBuffer & "<option value=""" & rs("ID") &""">" & rs("source") & "</option>"
				rs.MoveNext
			Wend
			d("from") = strBuffer

			sql = "SELECT * FROM blog_author ORDER BY ID ASC"
			strBuffer = ""
			Set rs = Db.Execute(sql)
			While Not rs.Eof
				strBuffer = strBuffer & "<option value=""" & rs("ID") &""">" & rs("author") & "</option>"
				rs.MoveNext
			Wend
			d("author") = strBuffer

			d("Tags") = ""
			sql = "SELECT * FROM blog_categories ORDER BY ID ASC"
			strBuffer = ""
			Set rs = Db.Execute(sql)
			While Not rs.Eof
				strBuffer = strBuffer & "<option value=""" & rs("ID") &""">" & rs("BlogCategoryName") & "</option>"
				rs.MoveNext
			Wend
			d("categories") = strBuffer
		End If

		d("message") = oFCKeditor.Create("txtContent")
		Set oFCKeditor = Nothing
		t.Load "submit.htm", d		
		Set rs = Nothing
	End Function
' ******************************************************************
	Function about()
		dim reval
		reval = reval & "<div class=""CNLTreeMenu"" id=""CNLTreeMenu2"">" & vbCrlf	
		reval = reval & BuildTree(0)
		reval = reval & "</div>" & vbCrlf
		d("tree") = reval
		t.Load "about.htm", d
	End Function
' ******************************************************************
	Function link()
		Set aaa = LoadModel("tag")
		
		Set aaa = Nothing
	End Function
' ******************************************************************
	Function getarchivesbymonth()
		Dim sql, tempBuffer
		tempBuffer = ""
		sql = "SELECT * FROM blog_blogs ORDER BY DateNode DESC"
		Set rs = Db.Execute(sql)
		While Not rs.Eof
			If tempBuffer <> CStr(rs("DateNode")) Then
				Response.Write rs("DateNode") & "<br>"
			End If
			tempBuffer = CStr(rs("DateNode"))
			Response.Write rs("Title") & "<br>" & rs("DateNode") & "<p>"
			rs.MoveNext
		Wend		
	End Function
' ******************************************************************
	' 获取验证码
	Function getcode()
		sb.Clear
		Response.Buffer=True
		Response.ExpiresAbsolute=Now()-1
		Response.CacheControl="no-cache"
		Response.Expires = -1
		Response.AddHeader "Pragma","no-cache"
		Response.AddHeader "cache-ctrol","no-cache"
		On Error Resume Next
		Dim zNum,i,j
		Dim Ados,Ados1
		Randomize Timer
		zNum = Cint(8999*Rnd+1000)
		Session("CommentCode") = zNum
		Dim zimg(4),NStr
		NStr=Cstr(zNum)
		For i=0 To 3
			zimg(i)=Cint(Mid(NStr,i+1,1))
		Next
		Dim Pos
		Set Ados=Server.CreateObject("Adodb.Stream")
		Ados.Mode=3
		Ados.Type=1
		Ados.Open
		Set Ados1=Server.CreateObject("Adodb.Stream")
		Ados1.Mode=3
		Ados1.Type=1
		Ados1.Open
		Ados.LoadFromFile(Server.mappath("body.Fix"))
		Ados1.write Ados.read(1280)
		For i=0 To 3
			Ados.Position=(9-zimg(i))*320
			Ados1.Position=i*320
			Ados1.write ados.read(320)
		Next	
		Ados.LoadFromFile(Server.mappath("head.fix"))
		Pos=lenb(Ados.read())
		Ados.Position=Pos
		For i=0 To 9 Step 1
			For j=0 To 3
				Ados1.Position=i*32+j*320
				Ados.Position=Pos+30*j+i*120
				Ados.Write Ados1.read(30)
			Next
		Next
		Response.ContentType = "image/BMP"
		Ados.Position=0
		Response.BinaryWrite Ados.read()
		Ados.Close:set Ados=nothing
		Ados1.Close:set Ados1=nothing
		If Err Then Session("CommentCode") = 9999
		Response.End
		
	End Function
End Class