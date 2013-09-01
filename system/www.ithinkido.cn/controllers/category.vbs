Class Category
	Function list()		
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

		t.Load "category/list.htm", d
	End Function


	Function BuildCategoryTree(parentId)

		Dim sql, rs, reval
		sql = "SELECT * FROM blog_categories WHERE ParentID = " & parentId & " ORDER BY Sort ASC"
		Set rs = Db.Execute(sql)
		reval = reval & "<ul>" & vbCrlf
		While Not rs.eof
			
			sql2 = "SELECT COUNT(*) AS NUM FROM blog_categories WHERE ParentID = " & rs("ID")  
			Set rs2 = Db.Execute(sql2)

			reval = reval & "<li class=""Child""><input type=""text"" value=""" & rs("BlogCategoryName") & """ id=""txtCategory" & rs("ID") & """><input type=""text"" value=""" &  rs("EName") & """ id=""txtEname" & rs("ID") & """><input type=""text"" id=""txtOrder" & rs("ID") & """ value=""" & rs("Sort") &""" /><input type=""button"" value=""更新"" onclick=""updatecategory(" &  rs("ID") & ")"">&nbsp;<a href=""index.asp?/category/list/up/" & rs("ID") & """> <img src=""system/img/up_1.gif""></a>&nbsp;<a href=""index.asp?/category/list/down/" & rs("ID") & """><img src=""system/img/down_1.gif""></a>&nbsp;<a  href=""index.asp?/category/list/del/" & rs("ID") & """>删除</a></li>" & vbCrlf

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