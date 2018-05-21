' Test CreateSuggestedPassword

Dim x
Set x = New Utility

AssertEqual Len(x.CreateSuggestedPassword), 32, ""

AssertEqual Len(x.CreateSuggestedPassword), 32, ""

AssertNotEqual x.CreateSuggestedPassword, x.CreateSuggestedPassword, ""

Set x = Nothing