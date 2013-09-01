Dim config
Set config = CreateObject("Scripting.Dictionary")

config("closed") = 0
config("isdebug") = 0
config("base_url") = "http://www.test.com/"
config("index_page") = "index.htm"
config("contact_email") = "wyrover@gmail.com"

config("sitetitle") = "Rover的Blog"
config("site_name") = "MySQL World"
config("site_description") = "blog测试"
config("webmaster") = "rover"
config("ICP") = "test"
config("siteurl") = "http://www.90bob.com"
config("sitelogo") = "img/xiao/logo.gif"
config("urlrewrite") = "0"


config("pathinfo") = 1


config("mail_server") = ""
config("mail_username") = ""
config("mail_password") = ""


config("scaffolding_trigger") = "scaffolding"


config("cookie_name") = "RoverApp"
config("cookie_name_setting") = "RoverAppSetting"
config("cookie_domain") = ".90bob.com"

config("skin") = "msdn"

' 存档文件夹
config("archives") = "\blog\"
config("attachment_img_dir") = "test"

config("rss_encoding") = "utf-8"
config("rss_language") = "zh-cn"


' 数据显示设置
config("article_list_count") = "10"
config("pic_list_count") = "20"
config("download_list_count") = "20"

' 是否允许站外提交
config("is_outsite_post") = "1"

' 允许上传的文件大小, 以K为单位
config("upload_filesize") = 50000000
config("upload_file_ext") = "rar|zip|gif|jpg|png"
config("upload_img_dir") = Server.MapPath(".") & config("archives") & "images\"

config("photo_custom") = "test$test2$test3$大师傅@@@ddddddd$eeeeeee"

config("photo_filesize") = 800
config("photo_file_ext") = "jpg|gif|bmp|png"
config("photo_is_rename") = 1

'自动分类 0不使用，1按年，2年-月，3年-月-日
config("photo_is_autosort") = 2			


config("PreviewType") = 999
config("PreviewImageWidth") = 120
config("PreviewImageHeight") = 100

config("TransitionColor") = "#ffffff"
config("database_name") = "rovercms.mdb"
config("database_connectionstring") = Server.MapPath(".") & "\system\" & Application_PATH & "\data\" & config("database_name")
config("database_backup_path") = Server.MapPath(".") & "\system\" & Application_PATH & "\data\"
config("count_file_path") = Server.MapPath(".") & "\system\" & Application_PATH & "\data\"


config("skins_path") = Server.MapPath(".") & "\skins\" 

config("custom_home_page") = "30"

'Response.Buffer = True
'Response.Cookies(config("cookie_name_setting")).Expires = Date + 365


Function ChkPost()

  Dim server_v1
  Dim server_v2  
  
  ChkPost=false  
  
  server_v1 = CStr(Request.ServerVariables("HTTP_REFERER"))
  server_v2 = Cstr(Request.ServerVariables("SERVER_NAME"))
  
  If Mid(server_v1,8,Len(server_v2))<>server_v2 Then
	chkpost=False
  Else
	chkpost=True  
  End If
   
End Function


Function GetBaseURL()
	GetBaseURL = config("base_url")
End Function

Function GetSiteName()
	GetSiteName = config("site_name")
End Function