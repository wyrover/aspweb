
Class Manage

	Function index()
		Call manage_check_login
		d("site") = Request.ServerVariables("http_host")
		t.Load "manage/manage.htm", d
	End Function 

	Function header()
		
		d("site") = Request.ServerVariables("http_host")
		t.Load "manage/header.htm", d
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

	Function BuildTree2(parentId)
	
		Dim sql, rs, reval
		sql = "SELECT * FROM t_menu WHERE ParentID = " & parentId & " ORDER BY Sort DESC"
		Set rs = Db.Execute(sql)
		While Not rs.eof			
			reval = reval & "d.add(" & rs("ID") & "," & rs("ParentID") & "," & "'" & rs("Name") & "','" &  rs("URL") &"');" & vbCrlf
			reval = reval & BuildTree2(rs("ID"))
			rs.MoveNext 
		Wend		
		Set rs = Nothing
		BuildTree2 = reval
		
	End Function

	Function sidebar()
		Call manage_check_login
		Dim navid, reval
		navid = Int(segment(3))
		
		reval = reval & "		d = new dTree('d');" & vbCrlf
		reval = reval & "d.config.target = ""mainFrame"";" & vbCrlf

		Select Case navid			
			Case 3
				reval = reval & "		d.add(0,-1,'我的站点', 'index.asp?/manage/desktop');" & vbCrlf
				reval = reval & BuildTree2(0)
				reval = reval & "		document.write(d);" & vbCrlf
				d("dtree") = reval			
		End Select

		t.Load "manage/sidebar.htm", d
	End Function

	Function desktop()
		Call manage_check_login
		If Request.Form("btnSave") <> "" Then
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

			regEx.Pattern = "^config\(\""base_url\""\).*\=.*?$"
			retVal = regEx.Replace(retVal, "config(""base_url"") = """ & Request.Form("txtBaseURL") & """")
			regEx.Pattern = "^config\(\""site_name\""\).*\=.*?$"
			retVal = regEx.Replace(retVal, "config(""site_name"") = """ & Request.Form("txtSitename") & """")
			regEx.Pattern = "^config\(\""site_description\""\).*\=.*?$"
			retVal = regEx.Replace(retVal, "config(""site_description"") = """ & Request.Form("txtSiteDescription") & """")
			regEx.Pattern = "^config\(\""webmaster\""\).*\=.*?$"
			retVal = regEx.Replace(retVal, "config(""webmaster"") = """ & Request.Form("txtWebmaster") & """")
			regEx.Pattern = "^config\(\""ICP\""\).*\=.*?$"
			retVal = regEx.Replace(retVal, "config(""ICP"") = """ & Request.Form("txtICP") & """")
			regEx.Pattern = "^config\(\""contact_email\""\).*\=.*?$"
			retVal = regEx.Replace(retVal, "config(""contact_email"") = """ & Request.Form("txtEmail") & """")
			regEx.Pattern = "^config\(\""archives\""\).*\=.*?$"
			retVal = regEx.Replace(retVal, "config(""archives"") = """ & Request.Form("txtArchives") & """")
			regEx.Pattern = "^config\(\""attachment_img_dir\""\).*\=.*?$"
			retVal = regEx.Replace(retVal, "config(""attachment_img_dir"") = """ & Request.Form("txtAttachmentImgDir") & """")

			' 数据显示设置
			regEx.Pattern = "^config\(\""article_list_count\""\).*\=.*?$"
			retVal = regEx.Replace(retVal, "config(""article_list_count"") = """ & Request.Form("txtArticleListCount") & """")
			regEx.Pattern = "^config\(\""pic_list_count\""\).*\=.*?$"
			retVal = regEx.Replace(retVal, "config(""pic_list_count"") = """ & Request.Form("txtPicListCount") & """")
			regEx.Pattern = "^config\(\""download_list_count\""\).*\=.*?$"
			retVal = regEx.Replace(retVal, "config(""download_list_count"") = """ & Request.Form("txtDownloadListCount") & """")
			
			Set regEx = Nothing
			
			Call SetCacheValue("<cms:function>GetBaseURL()</cms:function>", Request.Form("txtBaseURL"), 5)
			Call BuildFile(filepath, retVal, 0)
			Call MsgBox2("修改成功！",1,"index.asp?/manage/desktop")
		End If

		d("base_url") = config("base_url")
		d("site_name") = config("site_name")
		d("site_description") = config("site_description")
		d("webmaster") = config("webmaster")
		d("ICP") = config("ICP")
		d("contact_email") = config("contact_email")
		d("archives") = config("archives")
		d("attachment_img_dir") = config("attachment_img_dir")

		d("article_list_count") = config("article_list_count")
		d("pic_list_count") = config("pic_list_count")
		d("download_list_count") = config("download_list_count")

		d("server_name") = "http://" & Request.ServerVariables("server_name")
		d("sever_ip") =Request.ServerVariables("LOCAL_ADDR")
		d("server_engine") = ScriptEngine&"/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion
		d("server_software") = Request.ServerVariables("SERVER_SOFTWARE")
		d("server_port") = Request.ServerVariables("server_port")
		d("application_count") = Application.Contents.Count
		d("session_count") = Session.Contents.Count
		d("path") = Request.ServerVariables("path_translated")
		d("virtual_path") = "http://" & Request.ServerVariables("server_name") & Request.ServerVariables("script_name")
		d("server_timeout") = Server.ScriptTimeout
		d("server_time") = Now()

		d("obj1") = IsObj("Scripting.FileSystemObject")
		d("obj2") = IsObj("adodb.connection")
		d("obj3") = IsObj("Persits.Upload")
		d("obj4") = IsObj("Persits.Jpeg")
		d("obj5") = IsObj("Persits.MailSender")
		d("obj6") = IsObj("JMail.Message")
		d("obj7") = IsObj("CDONTS.NewMail")
		d("obj8") = IsObj("SoftArtisans.ImageGen")
		d("obj9") = IsObj("W3Image.Image")
		d("obj10") = IsObj("w3.Upload")


		Dim objFSO, objFile, objFolder
		Dim sb2, strPath
		Set sb2 = New StringBuilder
		strPath = config("skins_path")
		
		 
		' 创建一个文件系统对象
		Set objFSO = CreateObject("Scripting.FileSystemObject")

		
		For Each objFolder In objFSO.GetFolder(strPath).SubFolders				
			If FileExist(objFolder.Path & "\skin.png") Then			
				sb2.Append "<li><img src=""skins/" & objFolder.Name & "/skin.png"">" & objFolder.Name & "<input type=""button"" value=""选用此模板"" onclick=""setskin('" & objFolder.Name & "')""/></li>"
			End If
		Next	


		d("skins") = sb2.ToString()
		Set sb2 = Nothing



		t.Load "manage/desktop.htm", d
	End Function

	Function setskin()
		Response.Buffer = true
		Response.ExpiresAbsolute = Now - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"
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

		regEx.Pattern = "^config\(\""skin\""\).*\=.*?$"
		retVal = regEx.Replace(retVal, "config(""skin"") = """ & segment(3) & """")			
		
		Set regEx = Nothing		
	
		Call BuildFile(filepath, retVal, 0)

		Response.Write "1"
		Response.End
	End Function

	Function login()	
		Response.Buffer = True 
		Response.Expires = -1
		Response.ExpiresAbsolute = Now() - 1 
		Response.Expires = 0 
		Response.CacheControl = "no-cache" 

		Dim sql
		If Request.Form("txtUsername") <> "" Then
			Set rs = Db.CreateRS()
			sql = "SELECT * FROM member WHERE username = '" & Request.Form("txtUsername") & "'"
			rs.Open sql, Db.ConnectionString, 1, 2
			If Not rs.Eof Then
				If Md5(Request.Form("txtPassword"), 16) <> rs("Password") Then
					Call MsgBox2("您输入的密码不正确", 0, "0")				
				ElseIf CStr(session("GetCode")) <> CStr(Trim(request.Form("txtValidateCode"))) Then
					Call MsgBox2("您输入的确认码和系统产生的不一致,请重新输入.返回后请刷新登陆页面后重新输入正确的信息.", 0, "0")					
				End If

				
				Response.Cookies("username") = Request.Form("txtUsername")			
				Response.Cookies("password") = Request.Form("txtPassword")
				'Response.Cookies("username").Domain = config("cookie_domain")
				'Response.Cookies("password").Domain = config("cookie_domain")
				Response.Cookies("username").Expires = Date + 31
				Response.Cookies("password").Expires = Date + 31
				rs("logintimes") = rs("logintimes")+1
				rs("lastLoginIP") = GetClientIP()
				rs("lastLoginTime") = Now		
				rs.Update
				rs.Close
				Set rs = Nothing
				
			Else
				Call MsgBox2("不存在此用户名", 0, "0")				
			End If			
			
			Response.Redirect("index.asp?/manage/index")
		Else
			t.Load "manage/login.htm", d
		End If
	End Function 

	' 获取验证码
	Function getcode()
		Import "system/helpers/GetCode.vbs"
		sb.Clear		
		Randomize
		Call CreatValidCode("GetCode")
		Response.End		
	End Function
	
	Function logout()
		'Session.Abandon
		'Session(config("sitename") & "adminname") = ""		
		'Response.Redirect ("index.asp?/manage/login")

		Response.Cookies("username").Expires = Date - 1
		Response.Cookies("password").Expires = Date - 1
		Response.Write "<script>this.parent.location.href='index.asp?/manage/login';</script>"
	End Function



	Function blog_list()
		
		Dim pager, currentpage, classid
		Set pager = New Pager

		If segment(4) = "page" And segment(5) <> "" Then				
			currentpage = CInt(segment(5))						
		Else
			currentpage = 1
		End If

		dim searhtext
		searchtext = IIf(segment(3) <> "0", segment(3), "")

		Response.Cookies("ReturnURL") = "index.asp?/manage/blog_list/0/page/" & CStr(currentpage)

		dim sql
		sql = "select COUNT(ID) from blog_blogs WHERE Title LIKE '%%" & searchtext & "%%'"
		set rs = db.Execute(sql)
		Call pager.Init("index.asp?/manage/blog_list/" & searchtext & "/page/", currentpage, rs(0), 20, "")		
		set rs = Nothing			
		d("page1") = pager.getHTML()		
		d("pagesize") = pager.PageSize()
		If currentpage = 1 Then
			d("tableB") = "WHERE A.Title LIKE '%%" & searchtext & "%%'"
		Else
			d("tableB") = "WHERE A.Title LIKE '%%" & searchtext & "%%' AND A.ID NOT IN (SELECT TOP " & pager.PageSize() * (currentpage - 1) & " C.ID FROM blog_blogs C WHERE C.Title LIKE '%%" & searchtext & "%%' ORDER BY C.ID DESC)"
		End If	

		t.Load "manage/blog_list.htm", d
	End Function





	

	

	Function profile()
		Call manage_check_login
		t.Load "manage/profile.htm", d
	End Function

	Function removecache()
		Call manage_check_login
		Application.Contents.RemoveAll()
		Response.Write "缓存重载成功！"	
	End Function

	


	

	

	
	


	Function BuildCategoryTree(parentId)

		Dim sql, rs, reval
		sql = "SELECT * FROM blog_categories WHERE ParentID = " & parentId & " ORDER BY Sort ASC"
		Set rs = Db.Execute(sql)
		reval = reval & "<ul>" & vbCrlf
		While Not rs.eof
			
			sql2 = "SELECT COUNT(*) AS NUM FROM blog_categories WHERE ParentID = " & rs("ID")  
			Set rs2 = Db.Execute(sql2)

			reval = reval & "<li class=""Child""><input type=""text"" value=""" & rs("BlogCategoryName") & """ id=""txtCategory" & rs("ID") & """><input type=""text"" value=""" &  rs("EName") & """ id=""txtEname" & rs("ID") & """><input type=""text"" id=""txtOrder" & rs("ID") & """ value=""" & rs("Sort") &""" /><input type=""button"" value=""更新"" onclick=""updatecategory(" &  rs("ID") & ")"">&nbsp;<a href=""index.asp?/manage/category_list/up/" & rs("ID") & """> <img src=""images/manage/up_1.gif""></a>&nbsp;<a href=""index.asp?/manage/category_list/down/" & rs("ID") & """><img src=""images/manage/down_1.gif""></a>&nbsp;<a  href=""index.asp?/manage/category_list/del/" & rs("ID") & """>删除</a></li>" & vbCrlf

			If Not rs2.Eof AND Int(rs2("NUM")) > 0 Then			
				reval = reval & BuildCategoryTree(rs("ID"))		
			End If
			rs2.Close
			Set rs2 = Nothing
			rs.MoveNext 
		Wend		
		reval = reval & "</ul>" & vbCrlf
		Set rs = Nothing
		BuildCategoryTree = reval
	
	End Function


	Function category_list()
		Call manage_check_login
		Dim sql
		Dim nid,t1,t11,t2,t22
		Dim parentId

		sql = "SELECT ParentID FROM blog_categories WHERE ID = " & segment(4)
		Set rs = Db.Execute(sql)
		If Not rs.Eof Then
			parentId = rs("ParentID")
		End IF

		If segment(3) = "up" And segment(4) <> "0" Then
			sql = "SELECT ID, Sort FROM blog_categories WHERE ParentID = " & parentId & " ORDER BY Sort DESC, ID DESC"
			Set rs = Db.Execute(sql)
			Do While Not rs.Eof
				nid = Int(rs(0))
				If Int(segment(4)) = nid Then
					t22 = rs(1)
					rs.MoveNext
					If rs.Eof Then Exit Do
					t2 = rs(0)
					t11 = rs(1)
					Db.Execute "UPDATE blog_categories SET Sort = " & t11 & " WHERE ID = " & segment(4)
					Db.Execute "UPDATE blog_categories SET Sort = " & t22 & " WHERE ID = " & t2
					Exit Do
				End If
				rs.movenext
			Loop
		ElseIf segment(3) = "down" And segment(4) <> "0" Then
			sql = "SELECT ID, Sort FROM blog_categories WHERE ParentID = " & parentId & " ORDER BY Sort ASC, ID ASC"
			Set rs = Db.Execute(sql)
			Do While Not rs.Eof
				nid = Int(rs(0))
				If Int(segment(4)) = nid Then
					t22 = rs(1)
					rs.MoveNext
					If rs.Eof Then Exit Do
					t2 = rs(0)
					t11 = rs(1)
					Db.Execute "UPDATE blog_categories SET Sort = " & t11 & " WHERE ID = " & segment(4)
					Db.Execute "UPDATE blog_categories SET Sort = " & t22 & " WHERE ID = " & t2
					Exit Do
				End If
				rs.movenext
			Loop
		ElseIf segment(3) = "del" And segment(4) <> "0" Then			
			Db.Execute "DELETE FROM blog_blogs WHERE BlogCategoryID = " & segment(4)
			Db.Execute "DELETE FROM blog_categories WHERE ID = " & segment(4)
		End If

		
		d("category_list") = BuildCategoryTree(0)
		d("categories") = BuildCategoryTree2(0, "", 0)

		t.Load "manage/category_list.htm", d
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


	

	

	

	




	Function template()		
		If segment(3) = "1" Then
			d("templatelist") = getFileList(Server.MapPath(".") & "\system\" & Application_PATH & "\controllers\", "")
			d("dir") = "1"
			d("title") = "控制器管理"
		Else
			d("templatelist") = getFileList(Server.MapPath(".") & "\system\" & Application_PATH & "\views\", "")
			d("dir") = "0"
			d("title") = "视图管理"
		End If	
		t.Load "manage/template.htm", d
	End Function

	Function getFileList(path, prefix)
		dim fs, folder, file, item, url
		set fs = CreateObject("Scripting.FileSystemObject")
		set folder = fs.GetFolder(path)
		For Each item In folder.Files
			strBuffer = strBuffer & "<option value=""" & item.Path &""">" & prefix & "\" & item.Name & "</option>"
		Next	

		If folder.SubFolders.Count > 0 Then
			for each item in folder.SubFolders
			   strBuffer = strBuffer & getFileList(item.Path, item.Name)
			 next
		End If
		getFileList = strBuffer
	End Function

	Function getTemplateList()
		Response.Buffer = true
		Response.ExpiresAbsolute = Now - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"

		Dim dir
		If segment(3) = "1" Then
			dir = "\controllers\"
		Else
			dir = "\views\"
		End If

		Echo getFileList(Server.MapPath(".") & "\system\" & Application_PATH & dir, "")
	End Function

	Function readtemplate()
		Dim retVal, filepath
		retVal = ""
		filepath = IIf(Request.Form("file") <> "", Request.Form("file"), "")
		If filepath <> "" Then
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
		End If
		Echo retVal
	End Function

	Function createFile()
		Dim filepath, retVal, dir
		filepath = IIf(Request.Form("filepath") <> "", Request.Form("filepath"), "")
		retVal = IIf(Request.Form("content") <> "", Request.Form("content"), "")
		If segment(3) = "1" Then
			dir = "\controllers\"
		Else
			dir = "\views\"
		End If

		If filepath <> "" And retVal <> "" Then			
			filepath = Server.MapPath(".") & "\system\" & Application_PATH & dir & filepath
			Call BuildFile(filepath, retVal, 0)
			Response.Write "1"		
		End If
	End Function

	Function modifyFile()
		Dim filepath, retVal
		filepath = IIf(Request.Form("filepath") <> "", Request.Form("filepath"), "")
		retVal = IIf(Request.Form("content") <> "", Request.Form("content"), "")
		If filepath <> "" And retVal <> "" Then		
			Call BuildFile(filepath, retVal, 0)
			Response.Write "1"		
		End If
	End Function

	Function deleteFile()
		Dim filepath
		filepath = IIf(Request.Form("file") <> "", Request.Form("file"), "")
		If filepath <> "" Then
			Call DeleteFiles(filepath)
			Response.Write "1"
		End If
	End Function



	Function GetRomoteFile()
		Import "system/libraries/clsRemoteFile.vbs"
		Import "system/helpers/date_helper.vbs"

		Dim sContent
		Dim objRemoteFile
		sContent = Request.Form("fckeditor")

		Set objRemoteFile = New clsRemoteFile
		objRemoteFile.AllowExt = "GIF|JPG|PNG|BMP"		
		objRemoteFile.UploadDir = "attachments/" & config("attachment_img_dir") & "/"
		objRemoteFile.ContentPath =  config("base_url") & "attachments/" & config("attachment_img_dir") & "/"		

        Dim D_Name
        D_Name = "month_" & DateToStr(Now(), "ym")
		Call CreateFolder(objRemoteFile.UploadDir  & D_Name)        
			
		objRemoteFile.UploadDir = "attachments/" & config("attachment_img_dir") & "/" & D_Name & "/"
		objRemoteFile.ContentPath =  config("base_url") & "attachments/" & config("attachment_img_dir") & "/" & D_Name & "/"

		sContent = objRemoteFile.ReplaceRemoteUrl(sContent)
		Response.Write sContent
	End Function


End Class