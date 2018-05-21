' test GetAccount_VotingPower

Dim SteemIt
Set SteemIt = New Steem

Dim vp
vp = SteemIt.GetAccount_VotingPower("justyy")

AssertTrue vp >= 60 And vp <= 100, "justyy vp should be between 60 and 100"

Set SteemIt = Nothing