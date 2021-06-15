Function OpenExcel( _
		ByVal path As String, _
		Optional ByVal UpdateLinks As Boolean = False, _
		Optional ByVal ReadOnly As Boolean = False) As Workbook

	' Extract workbook name
	Dim wkbName As String
	wkbName = Right(path, InStrRev(path, "\") - 1)
	
	' Get workbook if open
	Dim wkb As Workbook
	On Error Resume Next
		Set wkb = Workbooks(wkbName)
	On Error Goto 0
	
	' Open workbook if close
	If wkb Is Nothing Then
		Set wkb = Workbooks.Open(path, UpdateLinks, ReadOnly)
	End If
		
	Set OpenExcel = wkb
End Function
	
