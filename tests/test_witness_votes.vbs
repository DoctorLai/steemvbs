' test GetAccount_WitnessVotes

Dim SteemIt
Set SteemIt = New Steem

Dim Util
Set Util = New Utility

Dim witness
witness = SteemIt.GetAccount_WitnessVotes("justyy")

AssertTrue Util.InArray("abit", witness), "justyy should vote abit"

AssertTrue Util.InArray("jerrybanfield", witness), "justyy should vote jerrybanfield"

Set SteemIt = Nothing
Set Util = Nothing