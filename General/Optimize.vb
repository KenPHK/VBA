Sub OptimizeOn()
	Application.Calculation = xlCalculationManual
	Application.ScreenUpdating = False
	Application.EnableEvents = False
End Sub

Sub OptimizeOff()
	Application.Calculation = xlCalculationAutomatic
	Application.ScreenUpdating = True
	Application.EnableEvents = True
End Sub
