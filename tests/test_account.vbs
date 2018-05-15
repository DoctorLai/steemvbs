' @justyy

Dim SteemIt
Set SteemIt = New Steem

Dim Account
Set Account = SteemIt.GetAccount("justyy")

' justyy should not be null
AssertNotNull Account, "Account Null"

' voting power should be positive
AssertTrue Account("voting_power") > 0, "voting power error"

' id should not change
AssertEqual Account("id"), 70955, "id not equal"

Dim a, b, c
a = SteemIt.GetAccount_Profile("justyy") ' cached
b = SteemIt.GetAccount_Profile("abit")   ' invalid cache
c = SteemIt.GetAccount_Profile("justyy") ' invalid cache

AssertTrue Len(a) > 0, ""

AssertTrue Len(b) > 0, ""

AssertTrue Len(c) > 0, ""

AssertEqual a, c, ""

Set Account = Nothing
Set SteemIt = Nothing