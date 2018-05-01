' Simple VBScript Unit Test
' @justyy

Option Explicit

Sub Assert(x, msg)
    If Not x Then
        Err.Raise 1, msg, msg              
    End If
End Sub

Sub AssertTrue(x, msg)
	Assert x, msg
End Sub

Sub AssertFalse(x, msg)
	Assert Not x, msg
End Sub

Sub AssertNull(x, msg)
	Assert IsNull(x), msg
End Sub

Sub AssertNotNull(x, msg)
	AssertFalse IsNull(x), msg
End Sub

Sub AssertEqual(x, y, msg)
	Assert x = y, msg
End Sub

Sub AssertNotEqual(x, y, msg)
	Assert x <> y, msg
End Sub
