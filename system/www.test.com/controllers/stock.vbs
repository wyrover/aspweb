Import "system/libraries/clsThief.vbs"	


Class stock


	Function index()		
		
		Dim letter, sbLetter
		
		Set sbLetter = New StringBuilder

		
			sbLetter.Append "<div class=""tabbertab"">"
			sbLetter.Append "<h2>新浪个股点评</h2>"
			'sbLetter.Append "<p>Tab " & Chr(letter) & " content.</p>"
			sbLetter.Append "</div>"


			sbLetter.Append "<div class=""tabbertab"">"
			sbLetter.Append "<h2>和讯股市直播</h2>"
			'sbLetter.Append "<p>Tab " & Chr(letter) & " content.</p>"
			sbLetter.Append "</div>"


			sbLetter.Append "<div class=""tabbertab"">"
			sbLetter.Append "<h2>上交所公告</h2>"
			'sbLetter.Append "<p>Tab " & Chr(letter) & " content.</p>"
			sbLetter.Append "</div>"


			sbLetter.Append "<div class=""tabbertab"">"
			sbLetter.Append "<h2>深交所公告</h2>"
			'sbLetter.Append "<p>Tab " & Chr(letter) & " content.</p>"
			sbLetter.Append "</div>"

			sbLetter.Append "<div class=""tabbertab"">"
			sbLetter.Append "<h2>上海公开交易信息</h2>"
			'sbLetter.Append "<p>Tab " & Chr(letter) & " content.</p>"
			sbLetter.Append "</div>"

			sbLetter.Append "<div class=""tabbertab"">"
			sbLetter.Append "<h2>深圳公开交易信息</h2>"
			'sbLetter.Append "<p>Tab " & Chr(letter) & " content.</p>"
			sbLetter.Append "</div>"


			sbLetter.Append "<div class=""tabbertab"">"
			sbLetter.Append "<h2>交易停复牌</h2>"
			'sbLetter.Append "<p>Tab " & Chr(letter) & " content.</p>"
			sbLetter.Append "</div>"


			sbLetter.Append "<div class=""tabbertab"">"
			sbLetter.Append "<h2>证券博客</h2>"
			'sbLetter.Append "<p>Tab " & Chr(letter) & " content.</p>"
			sbLetter.Append "</div>"

			sbLetter.Append "<div class=""tabbertab"">"
			sbLetter.Append "<h2>汇市指南</h2>"
			'sbLetter.Append "<p>Tab " & Chr(letter) & " content.</p>"
			sbLetter.Append "</div>"
		
		
		d("tabes") = sbLetter.ToString()
		Set sbLetter = Nothing
		t.Load "stock.htm", d
	End Function

	Function get_p()
		Response.Buffer = true
		Response.ExpiresAbsolute = Now - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"

		Dim matches, matches2
		Set objThief = New clsThief
		objThief.src = "http://www.1ting.com/group/group1_8.html"
		objThief.steal "utf-8"	

		Dim objRegex 
		Set objRegex = New RegExp
		objRegex.Global = True
		objRegex.IgnoreCase = True
		objRegex.MultiLine = True
		objRegex.Pattern = Replace("name=""@"">@</a>([^\b]+?)</ul>", "@", CStr(segment(3)))
		Set matches = objRegex.Execute(objThief.Value)

		If matches.Count > 0 Then		
		
			For Each m In matches				

				If m.SubMatches.Count > 0 Then		
				
					objRegex.Pattern = "href=""(.*?)""[^\b]+?>(.*?)</a>"
					Set matches2 = objRegex.Execute(m.SubMatches(0))
					If matches2.Count > 0 Then
						Echo "<ul>"
						For Each m2 In matches2
							If m2.SubMatches.Count > 0 Then
								Echo "<li><a href=""#"" onclick=""get_z_list('http://www.1ting.com" & m2.SubMatches(0) & "')"">" & m2.SubMatches(1) & "</a></li>"
							End If
						Next
						Echo "</ul>"
					End If
				
				End If		
				
			Next
		
			
		End If

	End Function


	Function getsina()			
		
		Set objThief = New clsThief
		objThief.src = "http://finance.sina.com.cn/column/ggdp.shtml"
		objThief.steal "gb2312"	

		Dim objRegex 
		Set objRegex = New RegExp
		objRegex.Global = True
		objRegex.IgnoreCase = True
		objRegex.MultiLine = True
		objRegex.Pattern = "<a\s+href=(.*?)\sclass=a06\s+target=_blank>(.*?)</a><font\s+class=z1>(.*?)</font>"
		Set matches = objRegex.Execute(objThief.Value)

		If matches.Count > 0 Then		
			Echo "<ul>"
			For Each m In matches				

				If m.SubMatches.Count > 0 Then		
					Echo "<li>"					
					Echo "<a href=""" & m.SubMatches(0) & """ target=_blank>" & m.SubMatches(1)	 & "</a>    <font color=red>" & m.SubMatches(2) & "</font>"					
					Echo "</li>"
				End If		
				
			Next
			Echo "</ul>"
			Echo "<div class=""clear""></div>"
			
		End If


	End Function


	Function gethexun()			
		
		Set objThief = New clsThief
		objThief.src = "http://stock.hexun.com/broadcast/"
		objThief.steal "gb2312"	

		Dim objRegex 
		Set objRegex = New RegExp
		objRegex.Global = True
		objRegex.IgnoreCase = True
		objRegex.MultiLine = True
		objRegex.Pattern = "<div\s+class=""sgd02"">(.*?)</div>[^\b]+?</span>[^\b]+?<a\s+href='(.*?)'\s+target=""_blank"">(.*?)</a>"
		Set matches = objRegex.Execute(objThief.Value)

		If matches.Count > 0 Then		
			Echo "<ul>"
			For Each m In matches				

				If m.SubMatches.Count > 0 Then		
					Echo "<li>"					
					Echo "<a href=""" & m.SubMatches(1) & """ target=_blank>" & m.SubMatches(2)	 & "</a>    <font color=red>" & m.SubMatches(0) & "</font>"					
					Echo "</li>"
				End If		
				
			Next
			Echo "</ul>"
			Echo "<div class=""clear""></div>"
			
		End If


	End Function


	Function getssedoc()			
		
		Set objThief = New clsThief
		objThief.src = "http://www.sse.com.cn/sseportal/ps/zhs/ggts/ssgsggqw_full.shtml"
		objThief.steal "gb2312"	

		Dim objRegex 
		Set objRegex = New RegExp
		objRegex.Global = True
		objRegex.IgnoreCase = True
		objRegex.MultiLine = True
		objRegex.Pattern = "<td\s+class=""content""\s+height=""22""><a\s+href=(.*?)\s+target=""_blank"">([^\b]+?)<span[^\b]+?>(.*?)</span>"
		Set matches = objRegex.Execute(objThief.Value)

		If matches.Count > 0 Then		
			Echo "<ul>"
			For Each m In matches				

				If m.SubMatches.Count > 0 Then		
					Echo "<li>"					
					Echo "<a href=""http://www.sse.com.cn" & m.SubMatches(0) & """ target=_blank>" & m.SubMatches(1)	 & "</a>    <font color=red>" & m.SubMatches(2) & "</font>"					
					Echo "</li>"
				End If		
				
			Next
			Echo "</ul>"
			Echo "<div class=""clear""></div>"
			
		End If


	End Function



	Function getsseTrade()			
		
		Set objThief = New clsThief
		objThief.src = "http://www.sse.com.cn/sseportal/webapp/datapresent/SSENewTradeInfoPublishAct"
		objThief.steal "gb2312"	

		Dim objRegex 
		Set objRegex = New RegExp
		objRegex.Global = True
		objRegex.IgnoreCase = True
		objRegex.MultiLine = True
		objRegex.Pattern = "<td\s+class=""content""\s+valign=""top""\s+width=""100%"">([^\b]+?)</td>"
		Set matches = objRegex.Execute(objThief.Value)

		If matches.Count > 0 Then		
			
			For Each m In matches				

				If m.SubMatches.Count > 0 Then		
								
					Echo m.SubMatches(0) 					
					
				End If		
				
			Next
			
			Echo "<div class=""clear""></div>"
			
		End If


	End Function


	Function getszsedoc()			
		
		'Set objThief = New clsThief
		'objThief.src = "http://www.sse.com.cn/sseportal/ps/zhs/ggts/ssgsggqw_full.shtml"
		'objThief.steal "gb2312"	

		'Dim objRegex 
		'Set objRegex = New RegExp
		'objRegex.Global = True
		'objRegex.IgnoreCase = True
		'objRegex.MultiLine = True
		'objRegex.Pattern = "尾页</a>[^\b]+Total:(.*?)</div>"
		'Set matches = objRegex.Execute(objThief.Value)

		'Dim pageCount

		'If matches.Count > 0 Then			
		'	For Each m In matches	
		'		If m.SubMatches.Count > 0 Then		
		'			pageCount = CInt(m.SubMatches(0))	
		'		End If				
		'	Next		
		'End If

		'For i <= pageCount
		'	objThief.src = "http://www.sse.com.cn/sseportal/ps/zhs/ggts/ssgsggqw_full.shtml"
		'	objThief.steal "gb2312"	
		'Next



	End Function


	Function getTecent()			
		
		Set objThief = New clsThief
		objThief.src = "http://stock.finance.qq.com/info/file/jiaoyi.shtml?979615"
		objThief.steal "gb2312"	

		Dim objRegex 
		Set objRegex = New RegExp
		objRegex.Global = True
		objRegex.IgnoreCase = True
		objRegex.MultiLine = True
		objRegex.Pattern = "(<table\s+width=""730""[^\b]+?class=""fontl26\s+marb6"">[^\b]+?)<!DOCTYPE"
		Set matches = objRegex.Execute(objThief.Value)

		If matches.Count > 0 Then		
			
			For Each m In matches				

				If m.SubMatches.Count > 0 Then		
								
					Echo Replace(Replace(m.SubMatches(0), "/info/file/tishi.htm", "http://stock.finance.qq.com/info/file/tishi.htm"), "/cgi-bin/info/weilaits", "http://stock.finance.qq.com/cgi-bin/info/weilaits") 					
					
				End If		
				
			Next
			
			Echo "<div class=""clear""></div>"
			
		End If


	End Function



	Function get_z_list()
		Dim src
		src = IIf(Request.Form("src") <> "", Request.Form("src"), "")			
		
		Set objThief = New clsThief
		objThief.src = src
		objThief.steal "utf-8"	

		Dim objRegex 
		Set objRegex = New RegExp
		objRegex.Global = True
		objRegex.IgnoreCase = True
		objRegex.MultiLine = True
		objRegex.Pattern = "<div\s+class=""zh"">[^\b]+?src=""(.*?)""[^\b]+?class=""zt""[^\b]+?href=""(.*?)""[^\b]+?>(.*?)</a>[^\b]+?</div>"
		Set matches = objRegex.Execute(objThief.Value)

		If matches.Count > 0 Then		
			Echo "<ul>"
			For Each m In matches				

				If m.SubMatches.Count > 0 Then		
					Echo "<li>"					
					Echo "<img src=""" & m.SubMatches(0) & """ alt=""" & m.SubMatches(2) & """ />"					 			 
					Echo "<a href=""#"" onclick=""get_wma_list('http://www.1ting.com" & m.SubMatches(1) & "')"">"
					Echo m.SubMatches(2) & "</a>"
					Echo "</li>"
				End If		
				
			Next
			Echo "</ul>"
			Echo "<div class=""clear""></div>"
			
		End If


	End Function

	Function getlinks()
		Echo "<a href=http://blog.sina.com.cn/hjhh/>花荣的blog</a>"
		Echo "<a href=http://blog.sina.com.cn/hjhh/>花荣的blog</a>"
		Echo "链接太多，对review是一种负担,以后还是减少del.icio.us和百度搜藏的使用"
	End Function



	Function getfx()			
		
		Set objThief = New clsThief
		objThief.src = "http://www2.fx168.com/fxnews/news_huishi.htm"
		objThief.steal "gb2312"	

		Dim objRegex 
		Set objRegex = New RegExp
		objRegex.Global = True
		objRegex.IgnoreCase = True
		objRegex.MultiLine = True
		objRegex.Pattern = "<p\s+class=""P10L"">(.*?)</p>[^\b]+?<p\s+class=""P10"">[^\b]+?<a\s+href='(.*?)'\s+[^\b]+?>([^\b]+?)</a>"
		Set matches = objRegex.Execute(objThief.Value)

		If matches.Count > 0 Then		
			Echo "<ul>"
			For Each m In matches				

				If m.SubMatches.Count > 0 Then		
					Echo "<li>"					
					Echo "<a href=""" & m.SubMatches(1) & """ target=_blank>" & m.SubMatches(2)	 & "</a>    <font color=red>" & m.SubMatches(0) & "</font>"					
					Echo "</li>"
				End If		
				
			Next
			Echo "</ul>"
			Echo "<div class=""clear""></div>"
			
		End If


	End Function

	Function get_wma_list()

		Dim src, sbwma
		Set sbwma = New StringBuilder
		src = IIf(Request.Form("src") <> "", Request.Form("src"), "")			
		
		Set objThief = New clsThief
		objThief.src = src
		objThief.steal "utf-8"	

		Dim objRegex 
		Set objRegex = New RegExp
		objRegex.Global = True
		objRegex.IgnoreCase = True
		objRegex.MultiLine = True
		objRegex.Pattern = "class=""ge"">[^\b]+?_(.*?)\.html"
		Set matches = objRegex.Execute(objThief.Value)
		
		If matches.Count > 0 Then		
		
			For Each m In matches				

				If m.SubMatches.Count > 0 Then		
					
					sbwma.Append "_"
					sbwma.Append m.SubMatches(0)
					
				End If		
				
			Next			
			

		End If			


		src = "http://play.1ting.com/p" & sbwma.ToString() & ".html"
		Set sbwma = Nothing
		objThief.src = src
		objThief.steal "utf-8"	

		objRegex.Global = True
		objRegex.IgnoreCase = True
		objRegex.MultiLine = True
		objRegex.Pattern = "type=""hidden""[^\b]+?name=""(.*?)""[^\b]+?value=""(.*?)""[^\b]+?_blank"">(.*?)</a>"
		Set matches = objRegex.Execute(objThief.Value)
		If matches.Count > 0 Then		
			Echo "<ul class=""list"">"
			For Each m In matches				
				
				If m.SubMatches.Count > 0 Then		
					Echo "<li><a href=""#"" onclick=""play('http://wma.1ting.com/wmam" & m.SubMatches(1) & "')"">" & m.SubMatches(0) & "</a>"					
				End If		
				
			Next			
			Echo "</ul>"
			Echo "<div class=""clear""></div>"

		End If		
		
		
	End Function


	Function weather()
		If GetCacheValue("weather") = "" Then
			Set objThief = New clsThief
			objThief.src = "http://Weather.love163.com/Site_Auto.Jsp?Purl=1"
			objThief.steal "gb2312"				
			Set objThief = Nothing
			Call SetCacheValue("weather", objThief.value, 1200)
		End If

		Response.Write GetCacheValue("weather")
	End Function

	Function test()
		If IsObjectInstalled("SoftArtisans.ImageGen")  = True Then
			Echo "安装成功"
		End If
	End Function

End Class