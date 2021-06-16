' Only work for text with length <= 255
Function SplitTextToArray( _
		ByVal text As String, _
		ByVal RowDelimiter As String, _
		ByVal ColDelimiter As String) As String
	
	' Replace row delimiter
	text = Replace(text, RowDelimiter, """;""")
	
	' Replace column delimiter
	text = Replace(text, ColDelimiter, """,""")
	
	' Insert Open and End bracket
	text = "{" & """" & text & """" & "}"
	
	' Convert text in an array format to an array
	SplitTextToArray = Evaluate(text)
End Function
