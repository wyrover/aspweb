
Class Cls_Cache 

	Public Reloadtime,MaxCount,CacheName
	Private LocalCacheName,CacheData,DelCount

	Private Sub Class_Initialize()
		Reloadtime=14400
		CacheName="Dvbbs"
	End Sub

	Private Sub SetCache(SetName,NewValue)
		Application.Lock
		Application(SetName) = NewValue
		Application.unLock
	End Sub

	Private Sub makeEmpty(SetName)
		Application.Lock
		Application(SetName) = Empty
		Application.unLock
	End Sub

	Public Property Let Name(ByVal vNewValue)
		LocalCacheName=LCase(vNewValue)
	End Property

	Public Property Let Value(ByVal vNewValue)
		If LocalCacheName<>"" Then
			CacheData=Application(CacheName&"_"&LocalCacheName)
			If IsArray(CacheData) Then
				CacheData(0)=vNewValue
				CacheData(1)=Now()
			Else
				ReDim CacheData(2)
				CacheData(0)=vNewValue
				CacheData(1)=Now()
			End If
				SetCache CacheName&"_"&LocalCacheName,CacheData
		Else
			Err.Raise vbObjectError + 1, "DvbbsCacheServer", " please change the CacheName."
		End If
	End Property

	Public Property Get Value()
		If LocalCacheName<>"" Then
			CacheData=Application(CacheName&"_"&LocalCacheName)
			If IsArray(CacheData) Then
				Value=CacheData(0)
			Else
				Err.Raise vbObjectError + 1, "DvbbsCacheServer", " The CacheData Is Empty."
			End If
		Else
			Err.Raise vbObjectError + 1, "DvbbsCacheServer", " please change the CacheName."
		End If
	End Property

	Public Function ObjIsEmpty()
		ObjIsEmpty=True
		CacheData=Application(CacheName&"_"&LocalCacheName)
		If Not IsArray(CacheData) Then Exit Function
		If Not IsDate(CacheData(1)) Then Exit Function
		If DateDiff("s",CDate(CacheData(1)),Now()) < 60*Reloadtime Then
			ObjIsEmpty=False
		End If
	End Function

	Public Sub DelCahe(MyCaheName)
		makeEmpty(CacheName&"_"&MyCaheName)
	End Sub

End Class



