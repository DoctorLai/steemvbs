' test GetAccount_Recovery

Dim SteemIt
Set SteemIt = New Steem

Dim re
re = SteemIt.GetAccount_Recovery("justyy")

AssertTrue re = "steem", "justyy's account recovery is not steem."

Set SteemIt = Nothing
