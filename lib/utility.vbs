' Utils

Option Explicit

Class Utility
	
	Public Function ValidateAccountName(value)
		Dim i, label, length, suffix
		suffix = "Account name should "
		
		If IsEmpty(value) Then
			ValidateAccountName = suffix + "not be empty."
			Exit Function
		End If 
		
		length = Len(value)
		If length < 3 Then
			ValidateAccountName = suffix + "be longer."
			Exit Function
		End If 
		
		If length > 16 Then
			ValidateAccountName = suffix + "be shorter."
			Exit Function
		End If 
		
		Dim Re 
		Set Re = New RegExp
		With Re
			.Pattern = "\."
			.IgnoreCase = False
			.Global = False
		End With
		
		If Re.Test(value) Then
			suffix = "Each account segment should "
		End If  
		
		Dim ref
		ref = Split(value, ".")
		
		length = UBound(ref)
		For i = 0 to length
			label = ref(i)
			
			Re.Pattern = "^[a-z]"
			If Not Re.test(label) Then
				ValidateAccountName = suffix + "start with a letter."
				Exit Function
			End If 
			
			Re.Pattern = "^[a-z0-9-]*$"
			If Not Re.test(label) Then
				ValidateAccountName = suffix + "have only letters, digits, or dashes."
				Exit Function
			End If 			
		
			Re.Pattern = "--"
			If Re.test(label) Then
				ValidateAccountName = suffix + "have only one dash in a row."
				Exit Function
			End If 	
			
			Re.Pattern = "[a-z0-9]$"
			If Not Re.test(label) Then
				ValidateAccountName = suffix + "end with a letter or digit."
				Exit Function
			End If 	
			
			If Not (Len(label) >= 3) Then
				ValidateAccountName = suffix + "be longer."
				Exit Function
			End If									
		Next
				
		ValidateAccountName = Empty
	End Function
	
	' Check value in Array
	Public Function InArray(needle, haystack)
		InArray = False
		needle = Trim(needle)
		Dim hay
		For Each hay in haystack
			If Trim(hay) = needle Then
				InArray = True
				Exit For
			End If
		Next
	End Function	
	
	' Get original URL from comment permlink
	Public Function GetUrlFromCommentPermLink(url)
		Dim Author, Link, Re
		Set Re = New RegExp
		With Re
			.Pattern = "(re-\w+-)*((\w+\-)*)"
			.Global = False
		End With				
		Dim my
		Set my = Re.Execute(url)
		If (my.Count >= 1) Then			
			Author = Split(my(0).submatches(0), "-")(1)
			Link = Mid(my(0).submatches(1), 1, Len(my(0).submatches(1)) - 1)
			GetUrlFromCommentPermLink = "https://steemit.com/@" + Author + "/" + Link
		Else
			GetUrlFromCommentPermLink = Empty
		End If
		Set Re = Nothing
	End Function	
	
	' Create suggested password
	Public Function CreateSuggestedPassword
		Const PASSWORD_LENGTH = 32
		Const AlphaBet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
		Randomize
		Dim i, s, t
		s = ""
		For i = 1 to PASSWORD_LENGTH
			t = Int(Rnd * Len(AlphaBet)) + 1
			s = s + Mid(AlphaBet, t, 1)
		Next
		CreateSuggestedPassword = s
	End Function
End Class 