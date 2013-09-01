Class Ajax
	Function login
		Dim username, password
		If Request.Form("username") <> "" Then				
			username = Request.Form("username")
			password = Request.Form("password")				
			
			Set rs = Db.CreateRS()
			sql = "select * from member where username = '" & username & "'"
			rs.Open sql, Db.ConnectionString, 1, 2
			If Not rs.Eof Then
					
				If Md5(password, 16) = rs("password") Then
					Response.Cookies("username") = Request.Form("username")			
					Response.Cookies("password") = password
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
					Response.Write "1"
				End If
			Else								
				Response.Write "0"
			End If
			
		End If
	End Function

	Function logout()
		'Response.Cookies("username").Domain = config("cookie_domain")
		'Response.Cookies("password").Domain = config("cookie_domain")
		Response.Cookies("username").Expires = Date - 1
		Response.Cookies("password").Expires = Date - 1
	End Function

	Function updatecategory()
		Set rs = Db.CreateRS()
		sql = "SELECT * FROM blog_categories WHERE ID = " & Request.Form("id")
		rs.Open sql, Db.ConnectionString, 1, 2
		If Not rs.Eof Then
			rs("BlogCategoryName") = Request.Form("name")
			rs("Ename") = Request.Form("ename")
			rs("Sort") = Request.Form("order")
			rs.Update
			rs.Close
			Response.Write "1"
		End If	
	End Function

	Function addcategory()
		Set rs = Db.CreateRS()
		sql = "SELECT * FROM blog_categories WHERE 1 > 2"
		rs.Open sql, Db.ConnectionString, 1, 2
		rs.AddNew
			rs("BlogCategoryName") = Request.Form("name")	
			rs("Ename") = Request.Form("ename")
			rs("ParentID") = Request.Form("parentId")
			rs("Sort") = Request.Form("order")			
		rs.Update
		rs.Close
		Response.Write "1"
	End Function

	Function updatelink()
		Set rs = Db.CreateRS()
		sql = "SELECT * FROM blog_links WHERE ID = " & Request.Form("id")
		rs.Open sql, Db.ConnectionString, 1, 2
		If Not rs.Eof Then
			rs("Name") = Request.Form("name")
			rs("URL") = Request.Form("url")			
			rs("Sort") = Request.Form("order")
			rs.Update
			rs.Close
			Response.Write "1"
		End If	
	End Function

	Function addlink()
		Set rs = Db.CreateRS()
		sql = "SELECT * FROM blog_links WHERE 1 > 2"
		rs.Open sql, Db.ConnectionString, 1, 2
		rs.AddNew
			rs("Name") = Request.Form("name")			
			rs("URL") = Request.Form("url")
			rs("Sort") = Request.Form("order")			
		rs.Update
		rs.Close
		Response.Write "1"
	End Function

	Function gethits()
		Response.Buffer=True
		Response.ExpiresAbsolute=Now()-1
		Response.CacheControl="no-cache"
		Response.Expires = -1
		Response.AddHeader "Pragma","no-cache"
		Response.AddHeader "cache-ctrol","no-cache"

		Dim id, fileid, fso, NewsFile, hits
		
		id = Int(segment(3))
		If id mod 1000 <> 0 Then
			fileid = Int((id / 1000) + 1)
		Else 
			fileid = Int(id / 1000)
		End If

		

		filepath = config("count_file_path") & fileid & ".txt"		

		'Response.Write filepath
		'Response.End

		If Not FileExist(filepath) Then
			Set fso = CreateObject("Scripting.FileSystemObject")
			Set NewsFile = fso.CreateTextFile(filepath, True)
			NewsFile.WriteLine("hits" & id & "=1") 					
			NewsFile.Close
			hits = 1
		Else
			Dim retVal		
				
			Set filestream = Server.CreateObject("ADODB.Stream")
			With filestream			
				.Type = 2 '以本模式读取
				.Mode = 3 
				.Charset = "gb2312"
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

			regEx.Pattern = "^hits" & id & "\=([^\b]+?)$"

			Set matches = regEx.Execute(retVal)
			If matches.Count > 0 Then		
				For Each m In matches	
					If m.SubMatches.Count > 0 Then		
						hits = m.SubMatches(0)		
					End If
				Next

				hits = hits + 1
				retVal = regEx.Replace(retVal, "hits" & id & "=" & hits)
				Call BuildFile(filepath, retVal, 1)
				
			Else
				Set fso = CreateObject("Scripting.FileSystemObject")
				Set NewsFile = fso.OpenTextFile(filepath, 8)
				NewsFile.WriteLine("hits" & id & "=1") 			
				NewsFile.Close
				hits = 1
			End If

			
			Set regEx = Nothing
			
		End If

		'Response.Write hits
		Response.Write "document.write(""" & hits & """);"
		Response.End
	End Function

	Function gethits2()
		Response.Buffer=True
		Response.ExpiresAbsolute=Now()-1
		Response.CacheControl="no-cache"
		Response.Expires = -1
		Response.AddHeader "Pragma","no-cache"
		Response.AddHeader "cache-ctrol","no-cache"

		Dim id, fileid, fso, NewsFile, hits
		
		id = Int(segment(3))
		If id mod 1000 <> 0 Then
			fileid = Int((id / 1000) + 1)
		Else 
			fileid = Int(id / 1000)
		End If

		

		filepath = config("count_file_path") & fileid & ".txt"		

		'Response.Write filepath
		'Response.End

		If Not FileExist(filepath) Then
			Set fso = CreateObject("Scripting.FileSystemObject")
			Set NewsFile = fso.CreateTextFile(filepath, True)
			NewsFile.WriteLine("hits" & id & "=1") 					
			NewsFile.Close
			hits = 1
		Else
			Dim retVal		
				
			Set filestream = Server.CreateObject("ADODB.Stream")
			With filestream			
				.Type = 2 '以本模式读取
				.Mode = 3 
				.Charset = "gb2312"
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

			regEx.Pattern = "^hits" & id & "\=([^\b]+?)$"

			Set matches = regEx.Execute(retVal)
			If matches.Count > 0 Then		
				For Each m In matches	
					If m.SubMatches.Count > 0 Then		
						hits = m.SubMatches(0)		
					End If
				Next				
			Else				
				hits = 0
			End If

			
			Set regEx = Nothing
			
		End If

		'Response.Write hits
		Response.Write "document.write(""" & hits & """);"
		Response.End
	End Function

	Function blog_add()
		Response.Buffer=True
		Response.ExpiresAbsolute=Now()-1
		Response.CacheControl="no-cache"
		Response.Expires = -1
		Response.AddHeader "Pragma","no-cache"
		Response.AddHeader "cache-ctrol","no-cache"

		If segment(3) <> "" And segment(3) <> "0" Then
				Set rs = Db.CreateRS()
				sql = "SELECT * FROM blog_blogs WHERE ID =" & segment(3)
				rs.Open sql, Db.ConnectionString, 1, 2
			
				'rs("BlogCategoryID") = CID(Request.Form("ddlCategories"))
				'strTitle = HTMLEncode(Trim(Request.Form("txtTitle")))
				'If strLength(strTitle) < 1 Then
				'	Call MsgBox2("标题字数不能为空", 0, "0")	
				'End If
				
				'rs("Title") = strTitle
				'rs("Author") = CID(Request.Form("ddlAuthor"))
				'rs("Source") = CID(Request.Form("ddlfrom"))
				'rs("IsShow") = CID(Request.Form("ddlAttributes"))
				'rs("IsTop") = IsTop
				'rs("Alias") = IIf(Request.Form("txtAlias") = "", "0", Request.Form("txtAlias"))   

				'tempTags = Split(CheckStr(Request.Form("txtTags"))," ")					 
				'Set mytag = LoadModel("tag")
				
				'post_taglist = ""
					
				'添加新的Tag
				'For Each post_tag In tempTags

					
					'If Len(Trim(post_tag))>0 Then
						'post_taglist = post_taglist & "{" & mytag.insert(CheckStr(trim(post_tag))) & "}"

					'End if
				'Next
				
				'Call mytag.Tags(2)
				'Set mytag = Nothing
				
				'rs("Tags") = post_taglist


				rs("Introduce") = HTMLEncode(Request.Form("txtIntroduce"))
				rs("Content") = HTMLEncode(Request.Form("txtContent"))
				rs("UpdateTime") = Now
				rs.Update
				rs.Close
				Set rs = Nothing
			
				Set myPageCreate = LoadModel("PageCreate")
				Call myPageCreate.createblogcontentbyid(segment(3))				
				Set myPageCreate = Nothing

				Response.Write "1"
				Response.End
			End If		
	End Function




End Class