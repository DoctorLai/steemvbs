' test vote worth functions

Dim SteemIt
Set SteemIt = New Steem

Dim Util
Set Util = New Utility

Dim fund
fund = SteemIt.GetRewardFund
AssertTrue fund > 0, "Rewards Pool should be larger than zero"

Dim esp
esp = SteemIt.Vests_To_Steem(SteemIt.GetAccountVests("justyy"))
AssertTrue esp > 1000, "justyy's ESP should be at least 1000"

Dim price
price = SteemIt.GetMedianPrice
AssertTrue price > 0, "median price should be larger than 0"

Dim upvote_value
upvote_value = SteemIt.GetAccount_UpvoteValue("justyy", 100, 100)
AssertTrue upvote_value > 0.1, "full vote value should be at least $0.1"

Dim current_upvote_value
current_upvote_value = SteemIt.GetAccount_UpvoteValue("justyy", 50, 20)
AssertEqualFloat upvote_value * 0.5 * 0.2, current_upvote_value, 0.1, "current upvote value calculation error"

Set SteemIt = Nothing
Set Util = Nothing