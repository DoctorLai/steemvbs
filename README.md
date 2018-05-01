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

# Unit Tests
Unit tests can be run via

```
cscript.exe /Nologo tests.wsf tests\test_account.vbs
```

# Roadmap
The features of Steem-Js and Steem-Python will be brought in. 

# Notice
This library is under development. Beware.