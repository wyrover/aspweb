Function HTMLCode(fString)
    If Not IsNull(fString) Then
    fString = Replace(fString, Chr(13), "")
    fString = Replace(fString, Chr(10) & Chr(10), "</P><P>")
    fString = Replace(fString, Chr(34), "")
    fString = Replace(fString, Chr(10), "<BR>")
    HTMLCode = fString
    End If
End Function
	
Function HTMLEncode(fString)
	If Not IsNull(fString) Then
	fString = Replace(fString, ">", ">")
	fString = Replace(fString, "<", "<")
	fString = Replace(fString, Chr(32), " ")
	fString = Replace(fString, Chr(9), " ")
	fString = Replace(fString, Chr(34), """")
	fString = Replace(fString, Chr(39), "'")
	fString = Replace(fString, Chr(13), "")
	fString = Replace(fString, Chr(10) & Chr(10), "</P><P> ")
	fString = Replace(fString, Chr(10), "<BR> ")
	HTMLEncode = fString
	End If
End Function


