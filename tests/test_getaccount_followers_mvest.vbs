' test GetAccount_FollowersMVest

Dim SteemIt
Set SteemIt = New Steem

Dim c1, c2
c1 = SteemIt.GetAccount_FollowersMVest("justyy")
AssertTrue c1 > 154101235.57696211338, "GetAccount_FollowersMVest > 154101235"

Set SteemIt = Nothing
