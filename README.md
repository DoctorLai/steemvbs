# steemvbs
SteemVBS is the first Steem Library written in VBScript. Yes, it is VBScript. ;)

# Examples
Class `Steem` is declared in `lib\steem.vbs` and you can do something like this

```
Dim SteemIt
Set SteemIt = (New Steem)("https://rpc.steemviz.com")

WScript.Echo SteemIt.Node
SteemIt.Node = "https://api.steemit.com"
WScript.Echo SteemIt.Node

Dim Account
Set Account = SteemIt.GetAccount("justyy")
WScript.Echo Account("voting_power")
```

To run the example:

```
cscript.exe /Nologo steem.wsf examples\account.vbs
```

## Formater Reputation
```
Dim Format
Set Format = New Formatter
Const EPSILON = 1e-3

AssertEqualFloat Format.Reputation(95832978796820), 69.833, EPSILON, "Format.Reputation 95832978796820"
```

## ValidateAccountName
```
Dim u
Set u = New Utility

AssertEqual u.ValidateAccountName("justyy"), "", ""
```

## Get Profile String
```
Dim SteemIt
Set SteemIt = New Steem

WScript.Echo SteemIt.GetAccount_Profile("justyy")
```

## Get Witness Votes
```
Dim SteemIt
Set SteemIt = New Steem

Dim Util
Set Util = New Utility

Dim witness
witness = SteemIt.GetAccount_WitnessVotes("justyy")

AssertTrue Util.InArray("abit", witness), "justyy should vote abit"
```

## Adding Real time Voting Power
```
' test GetAccount_VotingPower

Dim SteemIt
Set SteemIt = New Steem

Dim vp
vp = SteemIt.GetAccount_VotingPower("justyy")

AssertTrue vp >= 60 And vp <= 100, "justyy vp should be between 60 and 100"

Set SteemIt = Nothing
```

## Adding Account Effective Steem Power
```
' test GetAccount_EffectiveSteemPower

Dim SteemIt
Set SteemIt = New Steem

Dim esp
esp = SteemIt.GetAccount_EffectiveSteemPower("justyy")

WScript.Echo esp
AssertTrue esp >= 20000, "justyy esp should be larger than 20000"

Set SteemIt = Nothing
```

## Adding CreateSuggestedPassword
```
Dim x
Set x = New Utility

WScript.Echo x.CreateSuggestedPassword

Set x = Nothing
```

## GetUrlFromCommentPermLink
This function returns the steem post url given a comment url

```
' Test GetUrlFromCommentPermLink

Dim x
Set x = New Utility

AssertEqual x.GetUrlFromCommentPermLink("re-tvb-re-justyy-re-tvb-45qr3w-20171011t144205534z"), "https://steemit.com/@tvb/45qr3w", ""

AssertEqual x.GetUrlFromCommentPermLink("re-justyy-daily-quality-cn-posts-selected-and-rewarded-promo-cn-20180520t153728557z"), "https://steemit.com/@justyy/daily-quality-cn-posts-selected-and-rewarded-promo-cn", ""

Set x = Nothing
```

## Vests to Steem Power
```
Dim SteemIt
Set SteemIt = New Steem

WScript.Echo SteemIt.VestsToSp(1234234)

Set SteemIt = Nothing
```

## Invalidate Cache
```
Dim SteemIt
Set SteemIt = New Steem

' fresh
WScript.Echo SteemIt.GetAccount_VotingPower("justyy")
' cached
WScript.Echo SteemIt.GetAccount_VotingPower("justyy")
' do not use cache
SteemIt.Cache = False
WScript.Echo SteemIt.GetAccount_VotingPower("justyy")

Set SteemIt = Nothing
```

## Vests
```
Dim SteemIt
Set SteemIt = New Steem

WScript.Echo SteemIt.GetAccount_VestingShares("justyy")

Set SteemIt = Nothing
```

## Delegated Vests
```
Dim SteemIt
Set SteemIt = New Steem

WScript.Echo SteemIt.GetAccount_DelegatedVestingShares("justyy")

Set SteemIt = Nothing
```

## Received Vests
```
Dim SteemIt
Set SteemIt = New Steem

WScript.Echo SteemIt.GetAccount_ReceivedVestingShares("justyy")

Set SteemIt = Nothing
```

## GetAccount_Recovery
```
Dim SteemIt
Set SteemIt = New Steem

Dim re
re = SteemIt.GetAccount_Recovery("justyy")

AssertTrue re = "steem", "justyy's account recovery is not steem."

Set SteemIt = Nothing
```

## GetAccount_Followers
```
Dim SteemIt
Set SteemIt = New Steem

Dim Util
Set Util = New Utility

Dim followers
followers = SteemIt.GetAccount_Followers("justyy")

AssertTrue Util.InArray("ericet", followers), "ericet should follow justyy"

Set SteemIt = Nothing
Set Util = Nothing
```

## GetAccount_Following
```
' test GetAccount_Following

Dim SteemIt
Set SteemIt = New Steem

Dim Util
Set Util = New Utility

Dim followers
followers = SteemIt.GetAccount_Following("justyy")

AssertTrue Util.InArray("abit", followers), "justyy should follow abit"
AssertTrue Util.InArray("ericet", followers), "justyy should follow ericet"
```

## GetAccount_FollowingCount And GetAccount_FollowersCount
```
Dim SteemIt
Set SteemIt = New Steem

Dim c1, c2
c1 = SteemIt.GetAccount_FollowingCount("justyy")
c2 = SteemIt.GetAccount_FollowersCount("justyy")

AssertTrue c1 < c2, "GetAccount_FollowingCount < GetAccount_FollowersCount"
AssertTrue c1 > 100, "GetAccount_FollowingCount > 100"
AssertTrue c2 > 100, "GetAccount_FollowersCount > 100"
```

## GetAccount_FollowersMVest
```
Dim SteemIt
Set SteemIt = New Steem

Dim c1, c2
c1 = SteemIt.GetAccount_FollowersMVest("justyy")
AssertTrue c1 > 154101235.57696211338, "GetAccount_FollowersMVest > 154101235"
```

## Vests and Steem Conversion: Vests_To_Steem, Steem_Per_MVests and Steem_To_Vests
```
Dim SteemIt
Set SteemIt = New Steem

WScript.Echo SteemIt.Steem_Per_MVests
WScript.Echo SteemIt.Vests_To_Steem(1)
WScript.Echo SteemIt.Steem_To_Vests(1)
```

### GetRewardFund, GetMedianPrice, GetAccount_UpvoteValue
```
Dim SteemIt
Set SteemIt = New Steem

Dim fund
fund = SteemIt.GetRewardFund
AssertTrue fund > 0, "Rewards Pool should be larger than zero"

Dim esp
esp = SteemIt.Vests_To_Steem(SteemIt.GetAccountVests("justyy"))
AssertTrue esp > 1000, "justyy's ESP should be at least 1000"

Dim price
price = SteemIt.GetMedianPrice
AssertTrue price > 0, "median price should be larger than 0"

Dim upvote_value, vp, voting_weight
vp = 100
voting_weight = 100
upvote_value = SteemIt.GetAccount_UpvoteValue("justyy", vp, voting_weight)
AssertTrue upvote_value > 0.1, "full vote value should be at least $0.1"

Dim current_upvote_value
current_upvote_value = SteemIt.GetAccount_UpvoteValue("justyy", 50, 20)
AssertEqualFloat upvote_value * 0.5 * 0.2, current_upvote_value, 0.1, "current upvote value calculation error"

Set SteemIt = Nothing
```

# Unit Tests
Unit tests can be run via

```
cscript.exe /Nologo tests.wsf tests\test_account.vbs
```

or you can call `run_tests.cmd` to run all tests in the test folder `tests`.

![](https://github.com/DoctorLai/steemvbs/blob/master/run_tests.jpg?raw=true)

# Roadmap
The features of Steem-Js and Steem-Python will be brought in. 

# Notice
This library is under development. Beware.
