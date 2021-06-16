Function FindArrayDimension(ByVal arr As Variant) As Long
	On Error Goto Result
	Dim i As Long, temp As Long
	Do While True
		i = i + 1
		temp = UBound(arr, i)
	Loop
	On Error Goto 0
	
Result:
	FindArrayDimension = i - 1
End Function
