Dim SteemIt
Set SteemIt = New Steem

Dim r, rr
Set r = SteemIt.GetDynamicGlobalPeroperties("call", "[""database_api"",""get_current_median_history_price"",[]]", 1)
Set rr = r("result")
WScript.Echo Replace(rr("base"), " SBD", "") / Replace(rr("quote"), " STEEM", "")
