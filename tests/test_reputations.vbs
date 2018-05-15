' @justyy

Dim Format
Set Format = New Formatter

Const EPSILON = 1e-3

AssertEqualFloat Format.Reputation(95832978796820), 69.833, EPSILON, "Format.Reputation 95832978796820"

AssertEqualFloat Format.Reputation(10004392664120), 61.0017, EPSILON, "Format.Reputation 10004392664120"

AssertEqualFloat Format.Reputation(30999525306309), 65.42219, EPSILON, "Format.Reputation 30999525306309"

AssertEqualFloat Format.Reputation(-37765258368568), -16.193832, EPSILON, "Format.Reputation -37765258368568"

Set Format = Nothing