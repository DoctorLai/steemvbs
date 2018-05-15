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

# Unit Tests
Unit tests can be run via

```
cscript.exe /Nologo tests.wsf tests\test_account.vbs
```

or you can call `run_tests.cmd` to run all tests in the test folder `tests`.

# Roadmap
The features of Steem-Js and Steem-Python will be brought in. 

# Notice
This library is under development. Beware.