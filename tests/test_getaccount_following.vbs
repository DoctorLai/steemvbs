' test GetAccount_Following

Dim SteemIt
Set SteemIt = New Steem

Dim Util
Set Util = New Utility

Dim followers
followers = SteemIt.GetAccount_Following("justyy")

AssertTrue Util.InArray("abit", followers), "justyy should follow abit"

AssertTrue Util.InArray("ericet", followers), "justyy should follow ericet"

Set SteemIt = Nothing
Set Util = Nothing