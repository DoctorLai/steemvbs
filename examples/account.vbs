
Dim SteemIt
Set SteemIt = (New Steem)("https://rpc.steemviz.com")

WScript.Echo SteemIt.Node
SteemIt.Node = "https://api.steemit.com"
WScript.Echo SteemIt.Node

Dim Account
Set Account = SteemIt.GetAccount("justyy")
WScript.Echo Account("voting_power")

WScript.Echo SteemIt.GetAccount_Profile("justyy")

Dim witness
witness = SteemIt.GetAccount_WitnessVotes("justyy")

For i = 0 To UBound(witness)
	WScript.Echo "justyy votes for " & witness(i)
Next
