Public Enum BorderChoice
	xlAll
	xlOutside
	xlInside
	xlLeft
	xlRight
	xlTop
	xlBottom
End Enum

Sub SetCellBorder( _
		ByVal rngTarget As Range, _
		ByVal BorderType As BorderChoice, _
		Optional ByVal LineType As XlLineStyle = xlContinuous, _
		Optional ByVal LineWeight As XlBorderWeight = xlThin)
	
	With rngTarget
		.Borders.LineStyle = xlLineStyleNone
		
		Select Case BorderType
		Case xlAll
			.Borders.LineStyle = LineType
			.Borders.Weight = LineWeight
			
		Case xlOutside
			.BorderAround LineStyle:=LineType, Weight:=LineWeight
			
		Case xlInside
			.Borders(xlInsideHorizontal).LineStyle = LineType
			.Borders(xlInsideVertical).LineStyle = LineType
			.Borders(xlInsideHorizontal).Weight = LineWeight
			.Borders(xlInsideVertical).Weight = LineWeight
			
		Case xlLeft
			.Borders(xlEdgeLeft).LineStyle = LineType
			.Borders(xlEdgeLeft).Weight = LineWeight
			
		Case xlRight
			.Borders(xlEdgeRight).LineStyle = LineType
			.Borders(xlEdgeRight).Weight = LineWeight
			
		Case xlTop
			.Borders(xlEdgeTop).LineStyle = LineType
			.Borders(xlEdgeTop).Weight = LineWeight
			
		Case xlBottom
			.Borders(xlEdgeBottom).LineStyle = LineType
			.Borders(xlEdgeBottom).Weight = LineWeight
			
		End Select
	End With
End Sub
