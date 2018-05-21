' test GetAccount_EffectiveSteemPower

Dim SteemIt
Set SteemIt = New Steem

Dim esp
esp = SteemIt.GetAccount_EffectiveSteemPower("justyy")

WScript.Echo esp
AssertTrue esp >= 20000, "justyy esp should be larger than 20000"

Set SteemIt = Nothing