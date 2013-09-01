Class FileObj_Class

	Dim FormName, FileName, FilePath, FileSize, FileType, FileStart, FileExt
	
	Public Function SaveToFile(Path)
		On Error Resume Next
		Dim oFileStream
		Set oFileStream = CreateObject("ADODB.Stream")
		oFileStream.Type = 1
		oFileStream.Mode = 3
		oFileStream.Open
		UpFileStream.Position = FileStart
		UpFileStream.CopyTo oFileStream, FileSize
		oFileStream.SaveToFile Path, 2
		oFileStream.Close
		Set oFileStream = Nothing 
	End Function
	
	Public Function FileData
		UpFileStream.Position = FileStart
		FileData = UpFileStream.Read (FileSize)
	End Function
End Class
