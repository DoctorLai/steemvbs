' SteemVBS Project
' Steem Class in VBScript
' @justyy

Option Explicit

Const Version = "0.0.1"
Const DefaultSteemAPINode = "https://api.steemit.com"

Class Steem
	
	Private iNode
	Private ErrorMessage
	Private CachedAccountData
	Private CachedAccountData_SteemDB
	Private UseCache
		
	' class constructor with parameters
	Public Default Function Init(API_Node)
		Node = API_node
		Set Init = Me		
	End Function
	
	' get version
	Public Function GetVersion
		GetVersion = Version
	End Function

	' get error
	Public Function GetError()
		GetError = ErrorMessage
	End Function
	
	' get current api node
	Public Property Get Node
		Node = iNode
	End Property
	
	' set steem node
	Public Property Let Node(ByVal API_Node)
		iNode = API_Node
		If IsEmpty(iNode) Or iNode = "" Then
			iNode = DefaultSteemAPINode
		End If
	End Property
	
	' class constructor
	Private Sub Class_Initialize()
		Node = DefaultSteemAPINode
		ErrorMessage = ""
		Set CachedAccountData = Nothing
		CachedAccountData_SteemDB = Null
		Cache = True
	End Sub
	
	' should we use cache
	Public Property Let Cache(ByVal v)
		UseCache = v
	End Property
	
	' should we use cache
	Public Property Get Cache
		Cache = UseCache
	End Property
			
	' cached account
	Public Property Get CachedAccount
		CachedAccount = CachedAccountData
	End Property
	
	' cached account for steemdb
	Public Property Get CachedAccountSteemDB
		CachedAccountSteemDB = CachedAccountData_SteemDB
	End Property
	
	' class destructor
	Private Sub Class_Terminate()
	
	End Sub		
	
	' post to Steem Node via MSXML.ServerXMLHTTP
	Public Function Exec(ByVal Method, ByVal Paramers)
		' Error Handling
		On Error Resume Next
		
		Dim xmlhttp		
		Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
		
		' Indicate that page that will receive the request and the
		' type of request being submitted
		xmlhttp.open "POST", Node, False
		
		'handle errors
		If Err Then            
			ErrorMessage = Err.Description & " [0x" & Hex(Err.Number) & "]"
			Exec = Null
		Else
			xmlhttp.setRequestHeader "Content-Type", "application/json; charset=UTF-8"
					
			' Send the data
			xmlhttp.send "{""jsonrpc"":""2.0"",""method"":""" + Method + """,""params"":[[""" + Paramers + """]],""id"":""0""}"
			
			' Return JSON Text
			Exec = Trim(xmlhttp.responseText)
		End If		
		
		Set xmlhttp = Nothing	
		'disable error handling again
		On Error Goto 0        
	End Function
	
	' get account
	Public Function GetAccount(id)
		Dim r
		r = Trim(Exec("get_accounts", id))
		If IsNull(r) Then
			Set GetAccount = Nothing
		Else 
			Dim json
			Set json = New VbsJson
			Dim o
			Set o = json.Decode(r)
			If Not IsEmpty(o("result")) Then
				Set GetAccount = o("result")(0)
				Set CachedAccountData = o("result")(0)
			Else 
				Set GetAccount = Nothing
				Set CachedAccountData = Nothing
			End If 
			Set json = Nothing
			Set o = Nothing
		End If					
	End Function
	
	' get_dynamic_global_properties
	Public Function GetDynamicGlobalPeroperties()
		Dim r
		r = Trim(Exec("get_dynamic_global_properties", ""))
		If IsNull(r) Then
			Set GetDynamicGlobalPeroperties = Null
		Else 
			Dim json
			Set json = New VbsJson
			Dim o
			Set GetDynamicGlobalPeroperties = json.Decode(r)
		End If	
	End Function
	
	' check cache
	Private Function CacheAvailable(id)
		' disable cache manually
		If Cache = False Then
			CacheAvailable = False
			Exit Function
		End If
		If CachedAccountData is Nothing Then
			CacheAvailable = False
			Exit Function
		End If
		CacheAvailable = LCase(id) = LCase(CachedAccountData("name"))
	End Function
	
	' get witness votes
	Public Function GetAccount_WitnessVotes(id)
		Dim acc
		If CacheAvailable(id) Then			
			Set acc = CachedAccountData
		Else
			Set acc = GetAccount(id)			
		End If 
		GetAccount_WitnessVotes = acc("witness_votes")
	End Function
	
	' get profile text string
	Public Function GetAccount_Profile(id)
		Dim acc
		If CacheAvailable(id) Then			
			Set acc = CachedAccountData
		Else
			Set acc = GetAccount(id)
		End If 
		Dim json
		Set json = New VbsJson		
		Dim o		
		Set o = json.Decode(acc("json_metadata"))		
		If Not IsEmpty(o) Then 
			GetAccount_Profile = o("profile")("about")
		Else
			GetAccount_Profile = ""
		End If 
		Set o = Nothing
		Set json = Nothing
	End Function	
	
	' get voting power
	Public Function GetAccount_VotingPower(id)
		Dim acc
		If CacheAvailable(id) Then			
			Set acc = CachedAccountData
		Else
			Set acc = GetAccount(id)
		End If 	
		Dim vp, last_vote_time, sec
		vp = acc("voting_power")
		last_vote_time = Replace(acc("last_vote_time"), "T", " ")
		sec = DateDiff("s", last_vote_time, Now)
		Dim regen
		regen = sec * 10000 / 86400 / 5
		Dim total_vp
		total_vp = (vp + regen) / 100
		If total_vp >= 100 Then
			total_vp = 100
		End If 
		GetAccount_VotingPower = total_vp
	End Function		
	
	' get vesting shares
	Public Function GetAccount_VestingShares(id)
		Dim acc
		If CacheAvailable(id) Then			
			Set acc = CachedAccountData
		Else
			Set acc = GetAccount(id)
		End If 	
		GetAccount_VestingShares = Replace(acc("vesting_shares"), " VESTS", "")
	End Function
	
	' get delegated vesting shares
	Public Function GetAccount_DelegatedVestingShares(id)
		Dim acc
		If CacheAvailable(id) Then			
			Set acc = CachedAccountData
		Else
			Set acc = GetAccount(id)
		End If 	
		GetAccount_DelegatedVestingShares = Replace(acc("delegated_vesting_shares"), " VESTS", "")
	End Function	
	
	' get received_vesting_shares
	Public Function GetAccount_ReceivedVestingShares(id)
		Dim acc
		If CacheAvailable(id) Then			
			Set acc = CachedAccountData
		Else
			Set acc = GetAccount(id)
		End If 	
		GetAccount_ReceivedVestingShares = Replace(acc("received_vesting_shares"), " VESTS", "")
	End Function	
	
	' convert vests to sp
	Public Function VestsToSp(vests)
		'TODO, use real data
		VestsToSp = vests / 2038
	End Function
	
	' get effective sp
	Public Function GetAccount_EffectiveSteemPower(id)
		Dim vests, vests_plus, vests_minus
		Dim sp, sp_plus, sp_minus
		' account owns
		vests = GetAccount_VestingShares(id)
		' received
		vests_plus = GetAccount_ReceivedVestingShares(id)
		' delegated
		vests_minus = GetAccount_DelegatedVestingShares(id)
		' convert to steem power
		sp = VestsToSp(vests)
		sp_plus = VestsToSp(vests_plus)
		sp_minus = VestsToSp(vests_minus)
		' simple math
		GetAccount_EffectiveSteemPower = sp + sp_plus - sp_minus
	End Function
	
	' call api from steemdb
	Private Function Exec_SteemDB(ByVal method, ByVal parameters)
		Dim URL
		URL = "https://steemdb.com/api/" + Trim(method) + "?" + Trim(parameters)
		
		' Error Handling
		On Error Resume Next
		
		Dim xmlhttp		
		Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
		
		' Indicate that page that will receive the request and the
		' type of request being submitted
		xmlhttp.open "Get", URL, False
		
		'handle errors
		If Err Then            
			ErrorMessage = Err.Description & " [0x" & Hex(Err.Number) & "]"
			Exec_SteemDB = Null
		Else
			' call the api
			xmlhttp.send
			
			' Return JSON Text
			Exec_SteemDB = Trim(xmlhttp.responseText)
		End If		
		
		Set xmlhttp = Nothing	
		'disable error handling again
		On Error Goto 0 		
	End Function
	
	' check cache for steemdb
	Private Function CacheAvailableSteemDB(id)
		' disable cache manually
		If Cache = False Then
			CacheAvailableSteemDB = False
			Exit Function
		End If
		If IsNull(CachedAccountData_SteemDB) Then
			CacheAvailableSteemDB = False
			Exit Function
		End If
		CacheAvailableSteemDB = True
	End Function
		
	' get followers list
	Public Function GetAccount_Followers(ByVal id)
		Dim r
		If CacheAvailableSteemDB(id) Then
			r = CachedAccountData_SteemDB
		Else 
			r = Trim(Exec_SteemDB("accounts", "account=" + id))
		End If 
		If IsNull(r) Then
			Set GetAccount_Followers = Nothing
		Else 
			Dim json
			Set json = New VbsJson
			Dim o		
			o = json.Decode(r)
			If Not IsEmpty(o(0)("followers")) Then				
				GetAccount_Followers = o(0)("followers")
				CachedAccountData_SteemDB = r
			Else 
				GetAccount_Followers = Nothing
				Set CachedAccountData_SteemDB = Nothing
			End If 
			Set json = Nothing
			Set o = Nothing
		End If		
	End Function		
	
	' get following list
	Public Function GetAccount_Following(ByVal id)
		Dim r
		If CacheAvailableSteemDB(id) Then
			r = CachedAccountData_SteemDB
		Else 
			r = Trim(Exec_SteemDB("accounts", "account=" + id))
		End If 
		If IsNull(r) Then
			Set GetAccount_Following = Nothing
		Else 
			Dim json
			Set json = New VbsJson
			Dim o		
			o = json.Decode(r)
			If Not IsEmpty(o(0)("following")) Then				
				GetAccount_Following = o(0)("following")
				CachedAccountData_SteemDB = r
			Else 
				GetAccount_Following = Nothing
				Set CachedAccountData_SteemDB = Nothing
			End If 
			Set json = Nothing
			Set o = Nothing
		End If		
	End Function	
	
	' get following count
	Public Function GetAccount_FollowingCount(ByVal id)
		Dim r
		If CacheAvailableSteemDB(id) Then
			r = CachedAccountData_SteemDB
		Else 
			r = Trim(Exec_SteemDB("accounts", "account=" + id))
		End If 
		If IsNull(r) Then
			Set GetAccount_FollowingCount = Nothing
		Else 
			Dim json
			Set json = New VbsJson
			Dim o		
			o = json.Decode(r)
			If Not IsEmpty(o(0)("following_count")) Then				
				GetAccount_FollowingCount = o(0)("following_count")
				CachedAccountData_SteemDB = r
			Else 
				GetAccount_FollowingCount = Nothing
				Set CachedAccountData_SteemDB = Nothing
			End If 
			Set json = Nothing
			Set o = Nothing
		End If		
	End Function	
	
	' get followers count
	Public Function GetAccount_FollowersCount(ByVal id)
		Dim r
		If CacheAvailableSteemDB(id) Then
			r = CachedAccountData_SteemDB
		Else 
			r = Trim(Exec_SteemDB("accounts", "account=" + id))
		End If 		
		If IsNull(r) Then
			Set GetAccount_FollowersCount = Nothing
		Else 
			Dim json
			Set json = New VbsJson
			Dim o		
			o = json.Decode(r)
			If Not IsEmpty(o(0)("followers_count")) Then				
				GetAccount_FollowersCount = o(0)("followers_count")
				CachedAccountData_SteemDB = r
			Else 
				GetAccount_FollowersCount = Nothing
				Set CachedAccountData_SteemDB = Nothing
			End If 
			Set json = Nothing
			Set o = Nothing
		End If		
	End Function	
End Class
