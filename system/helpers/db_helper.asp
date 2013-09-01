<%
	Dim mobjConnection
	Set mobjConnection = Nothing
	Public Sub OpenConnection() 		
	    If mobjConnection Is Nothing Then
			Set mobjConnection = CreateObject("ADODB.Connection")
			mobjConnection.ConnectionString = "Provider=SQLOLEDB.1;Password=;Persist Security Info=True;User ID=sa;Initial Catalog=NorthWind;Data Source=localhost"
			mobjConnection.Open
		End If
	End Sub

	Public Sub CloseConnection() 
		If Not mobjConnection Is Nothing Then           
		    'If mobjConnection.State = adStateOpen Then        
		       mobjConnection.Close                     
		       Set mobjConnection = Nothing
		    'End If    
		End If		
	End Sub

	Public Function GetRecordset(sSQL)
		Dim objRs 
		Call OpenConnection()
		
		Set objRs = CreateObject("ADODB.Recordset")			
		
		objRs.CursorLocation = 3 'adUseClient
		objRs.CursorType     = 0 'adOpenForwardOnly
		objRs.LockType       = 4 'adLockBatchOptimistic
		objRs.Open  sSQL, mobjConnection                    		
		If objRs.RecordCount > 0 Then
			objRs.MoveFirst                                  
		End If
		Set objRs.ActiveConnection = Nothing					
		Set GetRecordset = objRs	
		Call CloseConnection()
	End Function
	
	
	Public Function GetPagedRecordSet(sSelect,sFrom, sWhere,sSort,iPageSize,iPageIndex,iNumRows)
		Dim sSQL
		Dim sSQLTemplate
		Dim S2
		Dim rs
		
		S2 = Replace(sSort," DESC"," _S_")
		S2 = Replace(S2," ASC"," DESC")
		S2 = Replace(S2," _S_"," ASC")
		
		sSQLTemplate = "SELECT * FROM (SELECT TOP {PS} * FROM  (SELECT TOP {PS2} {C} FROM {F} {W} ORDER BY {S1}) as T1 ORDER BY {S2}) AS T2 ORDER BY {S1}  "
		
		sSQLTemplate = Replace(sSQLTemplate,"{S2}",S2,1,1)
		sSQLTemplate = Replace(sSQLTemplate,"{S1}",sSort,1,2)
		sSQLTemplate = Replace(sSQLTemplate,"{C}",sSelect,1)
		sSQLTemplate = Replace(sSQLTemplate,"{F}",sFrom,1)
		
		If sWhere<>"" Then
			sSQLTemplate = Replace(sSQLTemplate,"{W}","WHERE " & sWhere,1)
		Else
			sSQLTemplate = Replace(sSQLTemplate,"{W}","",1)
		End If
		
		sSQLTemplate = Replace(sSQLTemplate,"{PS}",iPageSize,1)
		sSQLTemplate = Replace(sSQLTemplate,"{PS2}",iPageSize*(1 + iPageIndex),1)
		
		If iPageIndex = 0 And iNumRows=0 Then
			sSQL = "SELECT Count(*) FROM " & sFROM 
			If sWhere<>"" Then
				sSQL = sSQL & " WHERE " & sWHERE
			End If
			Set rs = GetRecordSet(sSQL)
			iNumRows = rs(0).Value
			Set rs = Nothing
		End If			
		
		Set GetPagedRecordSet = GetRecordSet(sSQLTemplate)
		'Response.Write sSQLTemplate
		
	End Function
%>