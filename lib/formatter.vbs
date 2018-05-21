' Formatter

Option Explicit

Class Formatter

	' Return Log() scale reputation
	Public Function Reputation(ByVal rep)
		Dim neg, v, reps
		rep = Int(rep)
		reps = CStr(rep)
		If (rep < 0) Then
			reps = Mid(reps, 1, Len(rep) - 1)
		End If 
		v = Log(Abs(rep) - 10)/Log(10) - 9
		If (rep < 0) Then
			v = -v
		End If
		Reputation = v * 9 + 25
	End Function	
	
End Class