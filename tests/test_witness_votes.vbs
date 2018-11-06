' test GetAccount_WitnessVotes

Dim SteemIt
Set SteemIt = New Steem

Dim Util
Set Util = New Utility

Dim witness
witness = SteemIt.GetAccount_WitnessVotes("steemfuckeos")

AssertTrue Util.InArray("abit", witness), "steemfuckeos should vote abit"

AssertTrue Util.InArray("oflyhigh", witness), "steemfuckeos should vote oflyhigh"

Set SteemIt = Nothing
Set Util = Nothing