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
