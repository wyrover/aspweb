'*************LCache_Class.asp****************
Class LCache_Class
	'-------------------------------------------------------------
	'磁盘缓存类 LCache Ver 1.0 Build 20060810
	'-------------------------------------------------------------
	'用途：
	'	缓存html代码
	'	
	'类成员：
	'	(属性)
	'		Version			版本信息，只读
	'		Name			缓存文件名称，绝对路径，例如C:\aaa.rst，只写
	'		IsAvailable		缓存文件是否可用，只读
	'	(方法)
	'		Add(file_tmp)	将记录集file_tmp保存到缓存文件中（缓存文件名称应先通过Name属性设定）
	'		Load()			将缓存内容载入程序
	'		Clear			清除缓存文件（缓存文件名称应先通过Name属性设定）
	'-------------------------------------------------------------
	'作者：Lukin
	'联系：mylukin@gmail.com
	'-------------------------------------------------------------
	'
	'==========================操作缓存类示例开始===========================
	'
	'	Dim LCache
	'	Set LCache = New LCache_Class    '建立缓存对象
	'	
	'	thePath = "Cache/"
	'	
	'	LCache.Name = Server.Mappath(thePath&server.urlencode(FileName)&".LCT")    '缓存文件物理路径，文件名（包括扩展名）可自行定义    '设置缓存类的Name属性
	'	
	'	If LCache.IsAvailable Then    '如果缓存可用
	'		LoadLCache = LCache.Load()        '则加载缓存文件到记录集中
	'	Else                        '否则
	'		LoadLCache = Content
	'		LCache.Add TheBody    '将记录集加入缓存
	'	End If
	'	
	'	Set LCache = Nothing    '释放缓存对象
	'==========================操作缓存类示例结束===========================
	Private pName		'缓存文件名称，绝对路径，例如C:\aaa.rst
	Private pFso		'fso对象
	Private pVersion	'版本
	Private pExpireHours '缓存多少小时后过期

	Public Property Get Version()
		Version = pVersion
	End Property

	Public Property Let Name(ByVal str_tmp)
		pName = str_tmp
	End Property

	Public Property Get IsAvailable()
		Dim pRndMinutes,pExpireMinutes,pFile,pFileLastModifyTime
		If (pFso.FileExists(pName)) Then
			Randomize
			pRndMinutes = Int(9 * Rnd) + 1	'随机数字，避免所有缓存同时过期
			pExpireMinutes = 60 * pExpireHours + pRndMinutes

			Set pFile = pFso.GetFile(pName)
				pFileLastModifyTime = pFile.DateLastModified
			Set pFile = Nothing
				
			If DateDiff("n",pFileLastModifyTime,Now()) >= pExpireMinutes Then
				IsAvailable = False
			Else
				IsAvailable = True
			End If
		Else
			IsAvailable = False
		End If
	End Property

	Public Sub Add(ByRef file_tmp)
		Call Clear()
		Set Fout = pFso.CreateTextFile(pName)
			Fout.Write file_tmp
			Fout.Close
		Set Fout = Nothing
	End Sub
	
	Public Function Load()
		Set FinFile = pFso.OpenTextFile(pName)
			If Not FinFile.atEndOfStream Then '先确定还没有到达结尾的位置
				Load = FinFile.ReadAll '读取整个文件的数据
			End If
			FinFile.Close
		Set FinFile = Nothing
	End Function

	
	Public Sub Clear()
		If (pFso.FileExists(pName)) Then
			pFso.DeleteFile pName,True
		End If
	End Sub
	
	Private Sub Class_Initialize()
		pVersion = "磁盘缓存类 LCache Ver 1.0 Build 20050628"
		pExpireHours = LExpireHours	'默认过期时间为1天
		Set pFso = Server.CreateObject("Scripting.FileSystemObject")
	End Sub

	Private Sub Class_Terminate()
		Set pFso = Nothing
	End Sub
End Class