
Class ccClsTemplate

  Private ccStrCode,ccStrStorage
  Private ccStrCacheCode
  Private ccBlnPublicCache,ccBlnPrivateCache
  Private ccStrName,ccStrCookieName
  Private ccStrSaveDirection,fileview,ccStrSaveFile,dir
  Private ccObjStream,ccObjFSO,ccStrFormat,ccIntObject,ccObjText,ccIntFormat

  Private Sub Class_Initialize
    ccStrName = "default"    '默认名称
    ccBlnPublicCache = False
    ccBlnPrivateCache = False
    fileview = "cache.htm"
    ccStrSaveFile = "save_cache.html"
    ccStrCookieName = "ccClass_Template"  'Application对象名前缀
    ccStrFormat = "UTF-8"    'UTF-8|ASCII|GB2312|BIG5
    ccIntFormat = -1
    ccIntObject = 1        '默认读取/保存模板组件 1:ADODB.Stream 2:FSO
    dir = Server.MapPath(".") & "\system\" & Application_PATH & "\views\"  '默认根路径
  End Sub

  Public Property Let Name(ccStrName_in)
    ccStrName = LCase(Trim(ccStrName_in))
  End Property

  Public Property Let Format(ccStrFormat_in)
    ccStrFormat = ccStrFormat_in
    If InStr(LCase(Trim(ccStrFormat_in)),"utf") > 0 Then
      ccIntFormat = -1
    Else
      ccIntFormat = 0
    End If
  End Property

  Public Property Let Object(ccStrObject_in)
    ccStrObject_in = LCase(Trim(ccStrObject_in))
    If InStr(ccStrObject_in,"fso") > 0 Then
      ccIntObject = 2
    Else
      ccIntObject = 1
    End If
  End Property

  Public Property Let PublicCache(ccBlnPublicCache_in)
    If ccBlnPublicCache_in = True Then
      ccBlnPublicCache = True
    Else
      ccBlnPublicCache = False
    End If
  End Property

  Public Property Let PrivateCache(ccBlnPrivateCache_in)
    If ccBlnPrivateCache_in = True Then
      ccBlnPrivateCache = True
    Else
      ccBlnPrivateCache = False
    End If
  End Property





  Public Property Let SaveDirection(ccStrSaveDirection_in)
    ccStrSaveDirection = ccStrSaveDirection_in
  End Property

  Public Property Let SaveFile(ccStrSaveFile_in)
    If ccStrSaveFile_in <> "" Then
      ccStrSaveFile = ccStrSaveFile_in
    End If
  End Property

  Public Property Get Code
    Code = ccStrCode
  End Property

  Public Property Get Storage
    Storage = ccStrStorage
  End Property

  Public Sub ClearCache
    Call ClearPrivateCache
    Call ClearPublicCache
  End Sub

  Public Sub ClearPrivateCache
    ccStrCacheCode = ""
  End Sub

  Public Sub ClearPublicCache
    Application(ccStrCookieName&ccStrName) = ""
  End Sub

  Public Sub ClearStorage
    ccStrStorage = ""
  End Sub

  Public Sub ClearCode
    ccStrCode = ""
  End Sub

  Public Sub SaveFront
    ccStrStorage = ccStrCode & ccStrStorage
  End Sub

  Public Sub SaveLast
    ccStrStorage = ccStrStorage & ccStrCode
  End Sub

  Public Sub SaveCode
    Call SaveToFile(1)
  End Sub

  Public Sub SaveStorage
    Call SaveToFile(2)
  End Sub

  Public Sub SetVar(ccStrTag_in,ccStrValue_in)
    ccStrCode = RePlace(ccStrCode,ccStrTag_in,ccStrValue_in)
  End Sub

  Private Sub SaveToFile(ccIntCode_in)
    Dim ccStrSaveCode
    If ccIntCode_in = 1 Then
      ccStrSaveCode = ccStrCode
    Else
      ccStrSaveCode = ccStrStorage
    End If
    If ccIntObject = 1 Then
      Set ccObjStream = Server.CreateObject("ADODB.Stream")
      With ccObjStream
        .Type = 2
        .Mode = 3
        .Open
        .Charset = ccStrFormat
        .Position = ccObjStream.Size
        .WriteText ccStrSaveCode
        .SaveToFile dir & ccStrSaveDirection & "\" & ccStrSaveFile,2
        .Close
      End With
      Set ccObjStream = Nothing
    Else
      Set ccObjFSO = CreateObject("Scripting.FileSystemObject")
      If ccObjFSO.FileExists(dir & ccStrSaveDirection & "\" & ccStrSaveFile) = True Then
        ccObjFSO.DeleteFile(dir & ccStrSaveDirection & "\" & ccStrSaveFile)
      End If
      Set ccObjText = ccObjFSO.OpenTextFile(dir & ccStrSaveDirection & "\" & ccStrSaveFile,2,True,ccIntFormat)
      ccObjText.Write ccStrSaveCode
      Set ccObjText = Nothing
      Set ccObjFSO = Nothing
    End If
    ccStrSaveCode = ""
  End Sub

Public Function Load(fileview, d)

	ccStrCode = ""

	If ccBlnPrivateCache = True Then
		If ccFncIsEmpty(ccStrCacheCode) = False Then
			ccStrCode = ccStrCacheCode
			Exit Function
		End If
	End If

	If ccBlnPublicCache = True Then
		If ccFncIsEmpty(Application(ccStrCookieName&ccStrName)) = False Then
			ccStrCode = Application(ccStrCookieName&ccStrName)
			Exit Function
		End If
	End If

	If ccIntObject = 1 Then
		Set ccObjStream = Server.CreateObject("ADODB.Stream")

		With ccObjStream
			.Type = 2
			.Mode = 3
			.Open
			.Charset = ccStrFormat
			.Position = ccObjStream.Size
			.LoadFromFile dir & fileview
			ccStrCode = .ReadText
			.Close
		End With

		Set ccObjStream = Nothing
	Else
		Set ccObjFSO = CreateObject("Scripting.FileSystemObject")

		If ccObjFSO.FileExists(dir & fileview) = True Then
			Set ccObjText = ccObjFSO.OpenTextFile(dir & fileview,1,False,ccIntFormat)
			ccStrCode = ccObjText.ReadAll
			Set ccObjText = Nothing
		End If

		Set ccObjFSO = Nothing
	End If

	For Each item In d
		ccStrCode = Replace(ccStrCode, "{$" & item & "}", d.Item(item))
    Next

	sb.Append ccStrCode

	Set sb2 = New StringBuilder
	sb2.Append ccStrCode
	Load = sb2.ToString()

	If ccBlnPrivateCache = True Then
	  ccStrCacheCode = ccStrCode
	End If

	If ccBlnPublicCache = True Then
	  Application(ccStrCookieName&ccStrName) = ccStrCode
	End If

End Function

End Class

Function ccFncIsEmpty(ByRef ccStrValue_in)
  If IsNull(ccStrValue_in) Or IsEmpty(ccStrValue_in) Or ccStrValue_in = "" Then
    ccFncIsEmpty = True
  Else
    ccFncIsEmpty = False
  End If
End Function




