Class Link
	Function list()
		Call manage_check_login
		Dim sql
		Dim nid,t1,t11,t2,t22

		If segment(3) = "up" And segment(4) <> "0" Then
			sql = "SELECT ID, Sort FROM blog_links ORDER BY Sort DESC, ID DESC"
			Set rs = Db.Execute(sql)
			Do While Not rs.Eof
				nid = Int(rs(0))
				If Int(segment(4)) = nid Then
					t22 = rs(1)
					rs.MoveNext
					If rs.Eof Then Exit Do
					t2 = rs(0)
					t11 = rs(1)
					Db.Execute "UPDATE blog_links SET Sort = " & t11 & " WHERE ID = " & segment(4)
					Db.Execute "UPDATE blog_links SET Sort = " & t22 & " WHERE ID = " & t2
					Exit Do
				End If
				rs.movenext
			Loop
		ElseIf segment(3) = "down" And segment(4) <> "0" Then
			sql = "SELECT ID, Sort FROM blog_links ORDER BY Sort ASC, ID ASC"
			Set rs = Db.Execute(sql)
			Do While Not rs.Eof
				nid = Int(rs(0))
				If Int(segment(4)) = nid Then
					t22 = rs(1)
					rs.MoveNext
					If rs.Eof Then Exit Do
					t2 = rs(0)
					t11 = rs(1)
					Db.Execute "UPDATE blog_links SET Sort = " & t11 & " WHERE ID = " & segment(4)
					Db.Execute "UPDATE blog_links SET Sort = " & t22 & " WHERE ID = " & t2
					Exit Do
				End If
				rs.movenext
			Loop
		ElseIf segment(3) = "del" And segment(4) <> "0" Then					
			Db.Execute "DELETE FROM blog_links WHERE ID = " & segment(4)
		End If

		t.Load "link/list.htm", d
	End Function
End Class 