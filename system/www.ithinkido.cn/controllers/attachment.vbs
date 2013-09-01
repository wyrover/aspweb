Class attachment

    Function index()
        t.Load "attachment/index.htm", d
    End Function

    Function upload()
        Import "system/helpers/date_helper.vbs"
        Import "system/libraries/Upload_File.vbs"
        

		Dim UploadDir 
		UploadDir = "attachments/" & config("attachment_img_dir") & "/"
        Dim D_Name, F_Name

        D_Name = "month_" & DateToStr(Now(), "ym")
		Call CreateFolder(UploadDir  & D_Name)        
		UploadDir = "attachments/" & config("attachment_img_dir") & "/" & D_Name & "/"

      

       


        Dim UP_FileSize
        UP_FileSize = 1024000000


        Dim FileUP
        Set FileUP = New Upload_File
        'FileUP.GetDate(-1)
        Dim F_File, F_Type
        Set F_File = FileUP.File("File")
        F_Name = randomStr(1) & Year(Now) & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now) & "." & F_File.FileExt
        F_Type = FixName(F_File.FileExt)
        If F_File.FileSize > Int(UP_FileSize) Then
            d("status") = "<div style=""padding:6px""><a href='index.asp?/attachment/index'>文件大小超出，请返回重新上传</a></div>"
        ElseIf IsvalidFile(UCase(F_Type)) = False Then
            d("status") = "<div style=""padding:6px""><a href='index.asp?/attachment/index'>文件格式非法，请返回重新上传</a></div>"
        Else
            F_File.SaveAs Server.MapPath(UploadDir &  F_Name)
            d("status") = "<script>addUploadItem('" & F_Type & "','" & UploadDir & F_Name&"',""0"")</script>" & "<div style=""padding:6px""><a href='index.asp?/attachment/index'>文件上传成功，请返回继续上传</a></div>"

        End If
        Set F_File = Nothing
        Set FileUP = Nothing

        t.Load "attachment/status.htm", d

    End Function

End Class
