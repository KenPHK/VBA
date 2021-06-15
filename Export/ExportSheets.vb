Sub ExportSheet( _
		ByVal SelectedSheet As Variant, _
		ByVal OutPath As String, _
		ByVal OutName As String, _
		Optional ByVal PasteValue As Boolean = True, _
		Optional ByVal BreakLink As Boolean = True, _
		Optional ByVal IgnoreWarning As Boolean = False)
	
	' Find the output file format from the output path
	Dim OutFormat As Long
	OutFormat = Right(OutPath, InStrRev(OutPath,".") - 1)
	
	' VB project cannot be exported to non-xlsm file
	If OutFormat <> "xlsm" And ThisWorkbook.HasVBProject And Not IgnoreWarning Then
		
		Dim response As Long
		response = MsgBox("This is VBA code in this workbook." & vbNewLine & _
			"If you proceed, the VBA code will be erased." & vbNewLine & vbNewLine & _
			"Do you wish to proceed?", vbYesNo, "Do you wish to proceed?")
		
		If response = vbNo Then Exit Sub
	End If
	
	' Multiple pages cannot be exported to txt or csv
	If UBound(SelectedSheet) > 1 And (OutFormat = "csv" Or OutFormat = "txt") Then
		MsgBox "Multiple worksheets cannot be exported to txt or csv." & vbNewLine & vbNewLine & _
			"Please export these worksheets separately."
	End If
	
	' Export worksheets to a new workbook
	Dim OutWkb As Workbook
	Dim wks As Variant
	For Each wks In SelectedSheet
			
		If OutWkb Is Nothing Then
			wks.Copy
			Set OutWkb = ActiveWorkbook
		Else
			wks.Copy After:=Outwkb.Sheets(OutWkb.Sheets.Count)
		End If
	Next wks
			
	' Paste value if specified
	If PasteValue Then
		For Each wks In OutWkb.Worksheets
			wks.Cells.Copy
			wks.Cells.PasteSpecial xlPasteValues
		Next wks
	End If
	
	' Break all external links if specified
	If BreakLink Then
		Dim ExtLinks As Variant
		ExtLinks = OutWkb.LinkSource(Type:=xlLinkTypeExcelLinks)
			
		On Error Resume Next
			Dim link As Variant
			For Each link In ExtLinks
				OutWks.BreakLink Name:=link, Type:=xlLinkTypeExcelLinks
			Next link
		On Error Goto 0
	End If

	' Save workbook
	OutWkb.SaveAs OutPath
	OutWkb.Close False, False
End Sub
