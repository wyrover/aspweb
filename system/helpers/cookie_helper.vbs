' ****************************************
'	设置Cookie
' ****************************************
Public Function SetCookie(key, val, exptime)
	Response.Cookies(key) = val
	Response.Cookies(key).Expires = exptime
End Function

' ****************************************
'	读取Cookie
' ****************************************
Public Function GetCookie(key)
	GetCookie = Request.Cookies(key)
End Function

