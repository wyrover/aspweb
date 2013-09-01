
'--------定义部份------------------
dim sql_injdata
SQL_injdata = "'|and|exec|insert|select|delete|update|count|*|%|chr|mid|master|truncate|char|declare|1=1|1=2|;"
SQL_inj = split(SQL_Injdata,"|")
'--------POST部份------------------
If Request.QueryString<>"" Then
	For Each SQL_Get In Request.QueryString
		For SQL_Data=0 To Ubound(SQL_inj)
		if instr(Request.QueryString(SQL_Get),Sql_Inj(Sql_DATA))>0 Then
			Response.Write "<Script Language=JavaScript>alert('系统提示你!\n\n请不要在参数中包含非法字符尝试注入!\n\n');window.location="&"'"&"index.asp"&"'"&";</Script>"
			Response.end
		end if
		next
	Next
End If
'--------GET部份-------------------
If Request.Form<>"" Then
	For Each Sql_Post In Request.Form
	For SQL_Data=0 To Ubound(SQL_inj)
		if instr(Request.Form(Sql_Post),Sql_Inj(Sql_DATA))>0 Then
			Response.Write "<Script Language=JavaScript>alert('系统提示你!\n\n请不要在参数中包含非法字符尝试注入!\n\n');window.location="&"'"&"index.asp"&"'"&";</Script>" 
			Response.end
		end if
		next
		next
end if
