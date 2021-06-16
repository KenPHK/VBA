Sub ForceRecalFormula(ByVal rngTarget As Range)
	rngTarget.Replace What:="=", Replacement:="=", _
		LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
		SearchFormat:=False, ReplaceFormat:=False
End Sub
