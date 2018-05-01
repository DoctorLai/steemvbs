' SteemVBS Project
' Steem Class in VBScript
' @justyy

Option Explicit

Const DefaultSteemAPINode = "https://api.steemit.com"

Class Steem
	
	Private iNode
	Private ErrorMessage
		
	' class constructor with parameters
	Public Default Function Init(API_Node)
		Node = API_node
		Set Init = Me		
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
	End Sub
	
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
			Exec = xmlhttp.responseText		
		End If		
		
		Set xmlhttp = Nothing	
		'disable error handling again
		On Error Goto 0        
	End Function
	
	' get account
	Public Function GetAccount(id)
		Dim r
		r = Trim(Exec("get_accounts", "justyy"))
		If r = Null Then
			Set GetAccount = Null
		Else 
			Dim json
			Set json = New VbsJson
			Dim o
			Set o = json.Decode(r)
			If Not IsEmpty(o("result")) Then
				Set GetAccount = o("result")(0)
			Else 
				Set GetAccount = Null
			End If 
		End If			
	End Function
	
	' get_dynamic_global_properties
	Public Function GetDynamicGlobalPeroperties()
		Dim r
		r = Trim(Exec("get_dynamic_global_properties", ""))
		If r = Null Then
			Set GetDynamicGlobalPeroperties = Null
		Else 
			Dim json
			Set json = New VbsJson
			Dim o
			Set GetDynamicGlobalPeroperties = json.Decode(r)
		End If	
	End Function
		
End Class
