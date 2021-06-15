Public Sub ZipFile( _
		ByVal FileToZipFullPath As String, _
		ByVal ZippedFileFullPath As String)
	
	' Create an empty zip file
	Open ZippedFileFullPath For Output As #1
	Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18,0)
	Close #1
	
	' Copy file/ folder into the zip file
	Dim ShellApp As New Shell32.Shell
	ShellApp.Namespace(ZippedFileFullPath).CopyHere ShellApp.Namespace(FileToZipFullPath).items
	
	' Pause macro until zipping is completed
	On Error Resume Next
		Do Until ShellApp.Namespace(ZippedFileFullPath).items.Count = ShellApp.Namespace(FileToZipFullPath).items.Count
			Application.Wait(Now + TimeValue("0:00:01")
		Loop
	On Error Goto 0
End Sub
