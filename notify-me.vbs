Declare_Period()


Sub Declare_Period()
	Dim period
	period = InputBox("Please input timeout", "Enter timeout", 60)
	If Not IsEmpty(period) Then
		If IsNumeric(period) Then
			Countdown(period)
		Else
			MsgBox ("Error number")
		 	Call Declare_Period
		End If
	End If	
End Sub

Function Countdown(period)
	WScript.Sleep 1000 * 60 * period
	MsgBox "wwwwwwwwwwwwwwwwwwwter"
	Call Declare_Period
End Function
