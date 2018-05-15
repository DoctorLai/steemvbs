' Utils

Option Explicit

Class Utility
	Function ValidateAccountName(value)
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
		With re
			.Pattern = "\."
			.IgnoreCase = False
			.Global = False
		End With
		
		If re.Test(value) Then
			suffix = "Each account segment should "
		End If  
		
		Dim ref
		ref = Split(value, ".")
		
		length = UBound(ref)
		For i = 0 to length
			label = ref(i)
			
			re.Pattern = "^[a-z]"
			If Not re.test(label) Then
				ValidateAccountName = suffix + "start with a letter."
				Exit Function
			End If 
			
			re.Pattern = "^[a-z0-9-]*$"
			If Not re.test(label) Then
				ValidateAccountName = suffix + "have only letters, digits, or dashes."
				Exit Function
			End If 			
		
			re.Pattern = "--"
			If re.test(label) Then
				ValidateAccountName = suffix + "have only one dash in a row."
				Exit Function
			End If 	
			
			re.Pattern = "[a-z0-9]$"
			If Not re.test(label) Then
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
	
	Public Function InArray(needle, haystack)
		InArray = False
		needle = Trim(needle)
		Dim hay
		For Each hay in haystack
			If trim(hay) = needle Then
				InArray = True
				Exit For
			End If
		Next
	End Function	
End Class 
