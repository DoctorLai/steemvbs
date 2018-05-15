' Simple VBScript Unit Test
' @justyy

Option Explicit

Sub Assert(x, msg)
    If Not x Then
        Err.Raise 1, msg
    End If
End Sub

Sub AssertTrue(x, msg)
	Assert x, "AssertTrue: " & msg
End Sub

Sub AssertFalse(x, msg)
	Assert Not x, "AssertFalse: " & msg
End Sub

Sub AssertNull(x, msg)
	Assert IsNull(x), "AssertNull: " & msg
End Sub

Sub AssertNotNull(x, msg)
	AssertFalse IsNull(x), "AssertNotNull: " & msg
End Sub

Sub AssertEqual(x, y, msg)
	Assert x = y, "AssertEqual: " & msg & ": " & x & ", " & y
End Sub

Sub AssertNotEqual(x, y, msg)
	Assert x <> y, "AssertNotEqual: " & msg & ": " & x & ", " & y
End Sub

Sub AssertEqualFloat(x, y, EPSILON, msg)
	Assert Abs(x - y) <= EPSILON, "AssertEqualFloat: " & msg & ": " & x & ", " & y & ", " & EPSILON
End Sub

Sub AssertNotEqualFloat(x, y, EPSILON, msg)
	Assert Abs(x - y) > EPSILON, "AssertNotEqualFloat: " & msg & ": " & x & ", " & y & ", " & EPSILON
End Sub
