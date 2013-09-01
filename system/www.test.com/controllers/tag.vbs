Class Tag
	Function list()		
		
		If segment(3) = "del" And segment(4) <> "0" Then					
			Db.Execute "DELETE FROM blog_tags WHERE tag_id = " & segment(4)
		End If

		t.Load "tag/list.htm", d
	End Function
End Class 