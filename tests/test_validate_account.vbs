' Test ValidateAccountName

Dim u
Set u = New Utility

AssertEqual u.ValidateAccountName("justyy"), "", ""

AssertEqual u.ValidateAccountName("justyy**"), "Account name should have only letters, digits, or dashes.", ""

AssertEqual u.ValidateAccountName("a    "), "Account name should have only letters, digits, or dashes.", ""

AssertEqual u.ValidateAccountName("12341234"), "Account name should start with a letter.", ""

AssertEqual u.ValidateAccountName("  askd f"), "Account name should start with a letter.", ""

AssertEqual u.ValidateAccountName("-"), "Account name should be longer.", ""

AssertEqual u.ValidateAccountName("-aasdfasdfasdfasdfasdfasdfasdfasdf"), "Account name should be shorter.", ""

Set u = Nothing