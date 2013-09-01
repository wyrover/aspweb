Import "system/libraries/clsThief.vbs"	


Class Icons


	Function index()		
		Response.Buffer = true
		Response.ExpiresAbsolute = Now - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"		

		If Request.Form("driver") <> "" Then
			Response.write "11111111"
			Response.End
		End If


		Echo "<ul>"
		Echo getFileList(Server.MapPath(".") & "\icons\", "")
		Echo "</ul>"
	End Function

	Function getFileList(path, prefix)
		dim fs, folder, file, item, url
		set fs = CreateObject("Scripting.FileSystemObject")
		set folder = fs.GetFolder(path)
		strBuffer = strBuffer & "<hr>"
		For Each item In folder.Files
			strBuffer = strBuffer & "<li><img src=""" & Replace(Mid(item.Path, 24), "\", "/") &""" /></li>" 
		Next	

		If folder.SubFolders.Count > 0 Then
			for each item in folder.SubFolders
			   strBuffer = strBuffer & getFileList(item.Path, item.Name)
			 next
		End If
		getFileList = strBuffer
	End Function

	

End Class
