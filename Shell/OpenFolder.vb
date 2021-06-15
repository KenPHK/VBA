Public Sub OpenFolder( _
		ByVal path As String)
	
	Call Shell("explorer.exe" & " " & path, vbNormalFocus)
End Sub
