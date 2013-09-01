Import "system/libraries/clsThief.vbs"	


Class Player


	Function index()		
		Dim letter, sbLetter
		
		Set sbLetter = New StringBuilder

		For letter = 65 To 90
			sbLetter.Append "<div class=""tabbertab"">"
			sbLetter.Append "<h2>" & Chr(letter) & "</h2>"
			'sbLetter.Append "<p>Tab " & Chr(letter) & " content.</p>"
			sbLetter.Append "</div>"
		Next
		
		d("tabes") = sbLetter.ToString()
		Set sbLetter = Nothing
		t.Load "player.htm", d
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