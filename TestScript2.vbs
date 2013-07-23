Option Explicit

' Setup the query parameters.
Dim strService, strTopic, strItem, strValue

strService = "Excel"
strTopic   = "[Book1]Sheet1"

WScript.Echo ""
WScript.Echo "Check Excel is running"
WScript.Echo "------------------------------------------------------------"

' Create a DDEClient object.
Dim oDDEClient
Set oDDEClient = CreateObject("DDECOMClient.DDEClient")

Dim oDDEConv
Set oDDEConv = oDDEClient.OpenConversation(strService, strTopic)

WScript.Echo "Conv: " & oDDEConv.Service & "|" & oDDEConv.Topic
WScript.Echo "Open? " & oDDEConv.IsOpen()

WScript.Echo ""
WScript.Echo "Poke a value into a cell via a conversation"
WScript.Echo "------------------------------------------------------------"

strItem  = "R2C2"
strValue = "Conversation Poke Test"

oDDEConv.PokeTextItem strItem, strValue

WScript.Echo "Val : " & oDDEConv.RequestTextItem(strItem)

WScript.Echo ""
WScript.Echo "Poke a value into a cell via the client"
WScript.Echo "------------------------------------------------------------"

strItem  = "R2C2"
strValue = "Client Poke Test"

oDDEClient.PokeTextItem strService, strTopic, strItem, strValue

WScript.Echo "Val : " & oDDEClient.RequestTextItem(strService, strTopic, strItem)

WScript.Echo ""
WScript.Echo "Maximise Excel via a conversation"
WScript.Echo "------------------------------------------------------------"

oDDEConv.ExecuteTextCommand "[App.Maximize()]"

WScript.Echo ""
WScript.Echo "Minimise Excel via the client"
WScript.Echo "------------------------------------------------------------"

oDDEClient.ExecuteTextCommand strService, strTopic, "[App.Minimize()]"

' Cleanup
Set oDDEClient  = nothing
Set oDDEConv    = nothing

WScript.Echo ""
WScript.Echo "------------------------------------------------------------"
WScript.Echo "Tests completed"
