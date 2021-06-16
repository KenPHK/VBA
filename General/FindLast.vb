Function FindLastRow(
		ByVal wks As Worksheet, _
		Optional ByVal colNum As Long = 1) As Long
	
	FindLastRow = wks.Cells(wks.Rows.Count, colNum).End(xlUp).Row
End Function

Function FindLastCol( _
	ByVal wks As Worksheet, _
	Optional ByVal rowNum As Long = 1) As Long

	FindLastCol = wks.Cells(rowNum, wks.Columns.Count).End(xlToLeft).Column
End Function

