' test GetAccount_Followers/Following/Count

Dim SteemIt
Set SteemIt = New Steem

Dim c1, c2
c1 = SteemIt.GetAccount_FollowingCount("justyy")
c2 = SteemIt.GetAccount_FollowersCount("justyy")

AssertTrue c1 < c2, "GetAccount_FollowingCount < GetAccount_FollowersCount"
AssertTrue c1 > 100, "GetAccount_FollowingCount > 100"
AssertTrue c2 > 100, "GetAccount_FollowersCount > 100"

Set SteemIt = Nothing
