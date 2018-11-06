' test Steem_Per_MVeests

Dim SteemIt
Set SteemIt = New Steem

Dim Util
Set Util = New Utility

Dim x
x = SteemIt.Steem_Per_MVests

Dim y
y = SteemIt.Vests_To_Steem(1)

Dim z
z = SteemIt.Steem_To_Vests(1)

AssertEqualFloat 1, y * z, 1e-3, "steem to vests * vests to steem should be clost to 1"

AssertEqualFloat y * 1e3, x, 1e-3, "steem per mvests / 1e3 should be vests_to_steem(1)"

AssertTrue z > 2000, "1 sp should be more than 2000 VESTS"

Set SteemIt = Nothing
Set Util = Nothing