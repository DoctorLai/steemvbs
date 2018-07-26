' test GetAccount_Followers

Dim SteemIt
Set SteemIt = New Steem

Dim Util
Set Util = New Utility

Dim followers
followers = SteemIt.GetAccount_Followers("justyy")

AssertTrue Util.InArray("ericet", followers), "ericet should follow justyy"

Set SteemIt = Nothing
Set Util = Nothing