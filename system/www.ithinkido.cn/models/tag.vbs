Class Tag

	

	Dim Arr_Tags
	
	Private Sub Class_Initialize()
		
		Arr_Tags =  Application(config("cookie_name") & "_blog_Tags")
		IF Not IsArray(Arr_Tags) Then Reload
		
	End Sub

	Private Sub Class_Terminate()

	End Sub
  
	Public Sub Reload		
		Tags(2) '更新Tag缓存
	End sub  

	
	Function Tags(ByVal action)	
		IF Not IsArray(Arr_Tags) or action=2 Then
			Dim log_TagsList
			Dim m_Arr_Tags
			Set rs = Db.Execute("SELECT tag_id, tag_name, tag_count FROM blog_tags")
			SQL_QUERY_NUM = SQL_QUERY_NUM + 1
			TempVar=""
			While Not rs.EOF				
				log_TagsList = log_TagsList & TempVar & rs("tag_id") & "||" & rs("tag_name") & "||" & rs("tag_count")
				TempVar=","
				rs.MoveNext
			Wend
			Set rs=Nothing
			m_Arr_Tags = Split(log_TagsList, ",")

			Application.Lock
			Application(config("cookie_name") & "_blog_Tags") = m_Arr_Tags
			Application.UnLock								 
		End If		

		Arr_Tags =  Application(config("cookie_name") & "_blog_Tags")		
	End Function
  
  
	Public Function insert(tagName) '插入标签,返回ID号
		
		If checkTag(tagName) Then
			
			Db.execute("update blog_tags set tag_count=tag_count+1 where tag_name='" & tagName & "'")
			insert=Db.execute("select top 1 tag_id from blog_tags where tag_name='" & tagName & "'")(0)
		Else
			Db.execute("insert into blog_tags (tag_name,tag_count) values ('" & tagName & "',1)")
			insert=Db.execute("select top 1 tag_id from blog_tags order by tag_id desc")(0)
		End If
	End Function
  
  
	Public Function remove(tagID) '清除标签
		If checkTagID(tagID) Then
		Db.execute("update blog_tags set tag_count=tag_count-1 where tag_id=" & tagID)
		End If
	End Function
  
  Public function filterHTML(str) '过滤标签
   	If isEmpty(str) Or isNull(str) Or len(str)=0 Then
        Exit Function
   		filterHTML=str
	 else
        dim log_Tag,log_TagItem
		For Each log_TagItem IN Arr_Tags
	   	    log_Tag=Split(log_TagItem,"||")
			str=replace(str,"{"&log_Tag(0)&"}","<a href=""default.asp?tag="&Server.URLEncode(log_Tag(1))&""">"&log_Tag(1)&"</a><a href=""http://technorati.com/tag/"&log_Tag(1)&""" rel=""tag"" style=""display:none"">"&log_Tag(1)&"</a> ")
		Next
		Dim re
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
        re.Pattern="\{(\d)\}"
      	str=re.Replace(str,"")
		filterHTML=str
	end if
  end function
  
	Public function filterEdit(str) '过滤标签进行编辑
		If isEmpty(str) Or isNull(str) Or len(str)=0 Then
			Exit Function
			filterEdit=str
		Else
			dim log_Tag,log_TagItem

			For Each log_TagItem IN Arr_Tags
				log_Tag=Split(log_TagItem,"||")
				str=replace(str,"{"&log_Tag(0)&"}",log_Tag(1)&",")
			Next

			Dim re
			Set re=new RegExp
			re.IgnoreCase =True
			re.Global=True
			re.Pattern="\{(\d)\}"
			str = re.Replace(str,"")
			filterEdit = Left(str, Len(str)-1)
		End If
	End Function
 
	Private Function checkTag(tagName) '检测是否存在此标签（根据名称）
		checkTag = False
		Dim log_Tag, log_TagItem
		
		
		For Each log_TagItem In Arr_Tags
			
			log_Tag=Split(log_TagItem,"||")
			If Lcase(log_Tag(1)) = Lcase(tagName) Then checkTag=true:exit function
		Next
	End function
  
  Private function checkTagID(tagID) '检测是否存在此标签（根据ID）
   checkTagID=false
   dim log_Tag,log_TagItem
	For Each log_TagItem IN Arr_Tags
		log_Tag=Split(log_TagItem,"||")
		if int(log_Tag(0))=int(tagID) then checkTagID=true:exit function
	Next
  end function
 
 Public function getTagID(tagName) '获得Tag的ID
   getTagID=0
   dim log_Tag,log_TagItem
	For Each log_TagItem IN Arr_Tags
		log_Tag=Split(log_TagItem,"||")
		if lcase(log_Tag(1))=lcase(ClearHTML(tagName)) then getTagID=log_Tag(0):exit function
	Next
 end function
  
 
End Class