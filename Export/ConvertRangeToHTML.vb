Sub ConvertRangeToHTML( _
		ByVal path As String, _
		ByVal rng As Range, _
		Optional VisibleOnly As Boolean = False)
	
	' Delete existing htm file
	On Error Resume Next
		Kill(path)
	On Error Goto 0
	
	' Remove hidden cell
	If VisibleOnly Then
		rng.SpecialCells(xlCellTypeVisible).Copy
		
		Dim wkb As Workbook
		Set wkb = Workbooks.Add
		
		With wkb.Sheets(1).Cells(1)
			.PasteSpecial Paste:=8
			.PasteSpecial xlPasteValues, , False, False
			.PasteSpecial xlPasteFormats, , False, False
			.Select
		End With
		
		Application.CutCopyMode = False
		
		Set rng = wkb.Sheets(1).Usedrange
	End If
	
	' Publish range in HTML format
	ActiveWorkbook.PublishObjects.Add( _
		SourceType:=xlSourceRange, _
		Filename:=path, _
		Sheet:= rng.Worksheet.Name, _
		Source:=rng.Address, _
		HTMLType:=xlHTMLStatic).Publish
	
	' Close temporary workbook if open
	On Error Resume Next
		Call wkb.Close(False, False)
	On Error Goto 0
End Sub
	
