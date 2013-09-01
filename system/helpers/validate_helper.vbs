' ****************************************
'	验证Email
' ****************************************
Function CheckEmail(strng)
    CheckEmail = false
    Dim regEx, Match
    Set regEx = New RegExp
    regEx.Pattern = "^\w+((-\w+)|(\.\w+))*\@[A-Za-z0-9]+((\.|-)[A-Za-z0-9]+)*\.[A-Za-z0-9]+$"
    regEx.IgnoreCase = True
    Set Match = regEx.Execute(strng)
    if match.count then CheckEmail= true
End Function

' ***************************************
'	验证用户名
' ***************************************
Public Function Check_UserName(str)
	Dim regEx
	Set regEx = New RegExp
	regEx.IgnoreCase = True
	regEx.Global = True
	regEx.MultiLine = True

	regEx.Pattern = "^[a-z0-9_]{2,20}$"
	Check_UserName = regEx.Test(str)
	Set regEx = Nothing
End Function


