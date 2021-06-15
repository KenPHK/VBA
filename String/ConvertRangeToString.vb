Public Function ConvertRangeToString(
		ByVal rngTarget As Range, _
		Optional ByVal RowDelimiter As String = ";", _
		Optional ByVal ColDelimiter As String = ",") As String
	
	' Create an empty string to store the output
	Const CELLLENGTH = 255
	
	Dim BufferSize As Long
	BufferSize = CELLLENGTH * rngTarget.Cells.count
	
	Dim strOut As String
	strOut = Space(BufferSize)
	
	' Store range into an array
	Dim arrData As Variant
	arrData = rngTarget.Value
	
	' Loop through all rows
	Dim i As Long, j As Long, length As Long
	For i = LBound(arrData,1) To UBound(arrData,1)

		' Insert row delimiter, skip first row
		If i > LBound(arrData,1) Then
			Mid(strOut, length+1, Len(RowDelimiter)) = RowDelimiter
			length = length + Len(RowDelimiter)
		End If

		' Loop through all columns in each row
		For j = LBound(arrData,2) To UBound(arrData,2)

			' Assign more space if not enough
			If length + Len(arrData(i,j)) + 2 > Len(strOut) Then
				strOut = strOut & Space(CLng(BufferSize / 4))
			End If

			' Insert column delimiter, skip first column
			If j < LBound(arrData,2) Then
				Mid(strOut, length+1, Len(ColDelimiter)) = ColDelimiter
				length = length + Len(ColDelimiter)
			End If

			' Concantenate string
			Mid(strOut, length+1, Len(arrData(i,j))) = arrData(i,j)
			length = length + Len(arrData(i,j))
		Next j
	Next i

	ConvertRangeToString = Left(strOut, length) & RowDelimiter
End Function
