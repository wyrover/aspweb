

Class comment

	Function list()		
		Dim pager, currentpage, classid
		Set pager = New Pager

		If segment(3) = "page" And segment(4) <> "" Then				
			currentpage = CInt(segment(4))						
		Else
			currentpage = 1
		End If

		Response.Cookies("ReturnURL") = "index.asp?/comment/list/page/" & CStr(currentpage)

		dim sql
		sql = "select COUNT(ID) from blog_comments" 
		set rs = db.Execute(sql)
		Call pager.Init("index.asp?/comment/list/page/", currentpage, rs(0), 20, "")
		set rs = Nothing			
		d("page1") = pager.getHTML()		
		d("pagesize") = pager.PageSize()
		If currentpage = 1 Then
			d("tableB") = ""
		Else
			d("tableB") = "WHERE A.ID NOT IN (SELECT TOP " & pager.PageSize() * (currentpage - 1) & " C.ID FROM blog_comments C ORDER BY C.ID DESC)"	
		End If	
		t.Load "comment/list.htm", d
	End Function

	Function del()
		Call manage_check_login
		Dim returnURL, sql
		If Request.Cookies("ReturnURL") <> "" Then
			returnURL = Request.Cookies("ReturnURL")
		Else
			returnURL = "index.asp?/comment/list"
		End If
		sql = "DELETE FROM blog_comments WHERE ID IN (" & Request.Form("ID") & ")"
		Db.Execute(sql)
		Call MsgBox2("删除成功！", 1, returnURL)
	End Function


	Function post()		

		Dim blogId			
		Dim strUsername
		Dim strContent

		blogId =  segment(3)
		d("blogId") = segment(3)

		If Request.Form("comment") <> "" Then

			If Len(request.Form("txtValidateCode")) = 0 Or Trim(request.Form("txtValidateCode")) <> "3" Then
				Response.Write("您输入的确认码和系统产生的不一致,请重新输入.")	
				Response.End
			End If

			strUsername = Trim(Request.Form("txtAuthor"))
			If strLength(strUsername) = 0 Then
				Response.Write("名称或邮箱不能为空")	
				Response.End				
			End If


			strContent = HTMLEncode(Trim(Request.Form("comment")))
			If strLength(strContent) = 0 Or strLength(strContent) > 1000 Then
					Response.Write("留言不能为空或过长")	
					Response.End								
			End If

			'If CStr(session("CommentCode")) <> CStr(Trim(request.Form("txtValidateCode")))  Then
			'		Call MsgBox2("您输入的确认码和系统产生的不一致,请重新输入.返回后请刷新页面后重新输入正确的信息.", 0, "0")
			'ENd If		


			Dim commentId
			Dim posttime
			posttime = Now
			blogId =  segment(3)
			Set rs = Db.CreateRS()
			sql = "SELECT * FROM blog_comments WHERE 1 > 2"
			rs.Open sql, Db.ConnectionString, 1, 2		
			rs.AddNew
				rs("BlogID") = blogId		
				rs("Content") = strContent
				rs("Author") = strUsername			
				rs("PostTime") = posttime
				rs("PostIP") = GetClientIP()
				rs("Email") = Request.Form("txtEmail")
				rs("WebSite") = Request.Form("txtUrl")
			rs.Update
				commentId = rs("id")
			rs.Close

			Set rs = Nothing

			Dim myPageCreate
			Set myPageCreate = LoadModel("PageCreate")	
			Call myPageCreate.createblogcontentbyid(blogId)			
			Set myPageCreate = Nothing	

			Dim commentcount
			sql = "SELECT COUNT(ID) FROM blog_comments WHERE blogid = " & blogId
			Set rs = Db.Execute(sql)
			commentcount = rs(0)
			rs.Close
			Set rs = Nothing
			

			Response.Write("<ol class=""commentlist"">")
			Response.Write("<li class=""altcomment"">")
			Response.Write("<h3 class=""commenttitle""><a name=""cmt" & commentcount  & """>" & commentcount  & "</a>.<a href=""mailto:youremail@email.com"">" & strUsername & "</a></h3>")
			Response.Write(strContent)
			Response.Write("<p class=""commentmeta"">" & posttime & "</p>")
			Response.Write("</li>")
			Response.Write("</ol>")
			
			Response.End				
			
		End If

			
	End Function




End Class



