Option Explicit

' Setup the query parameters.
Dim strService, strTopic, strItem, strValue

strService = "PROGMAN"
strTopic   = "PROGMAN"
strItem    = "Accessories"

WScript.Echo ""
WScript.Echo "Create DDEClient object"
WScript.Echo "------------------------------------------------------------"

' Create a DDEClient object.
Dim oDDEClient
Set oDDEClient = CreateObject("DDECOMClient.DDEClient")

WScript.Echo ""
WScript.Echo "Request text item via client object"
WScript.Echo "------------------------------------------------------------"

' Request an item value, creating a conversation on-the-fly.
strValue = oDDEClient.RequestTextItem(strService, strTopic, strItem)

WScript.Echo "Link: "  & strService & "|" & strTopic & "!" & strItem
WScript.Echo "Val : " & strValue

WScript.Echo ""
WScript.Echo "Create a blank conversation, open it and request an item"
WScript.Echo "------------------------------------------------------------"

' Create a blank DDE conversation.
Dim oDDEConv
Set oDDEConv = CreateObject("DDECOMClient.DDEConversation")

With oDDEConv
	.Service = strService
	.Topic   = strTopic
End With

' Open it and request an item.
oDDEConv.Open()

WScript.Echo "Conv: " & oDDEConv.Service & "|" & oDDEConv.Topic
WScript.Echo "Open? " & oDDEConv.IsOpen()
WScript.Echo "Val : " & oDDEConv.RequestTextItem(strItem)

' Cleanup.
oDDEConv.Close()

WScript.Echo ""
WScript.Echo "Open multiple conversations to different services"
WScript.Echo "------------------------------------------------------------"

Dim oDDEConv1, oDDEConv2, oDDEConv3

' Open more conversations via the DDEClient object.
Set oDDEConv1 = oDDEClient.OpenConversation("PROGMAN", "PROGMAN")
Set oDDEConv2 = oDDEClient.OpenConversation("Folders", "AppProperties")
Set oDDEConv3 = oDDEClient.OpenConversation("Shell",   "AppProperties")

WScript.Echo "Conv: " & oDDEConv1.Service & "|" & oDDEConv1.Topic
WScript.Echo "Open? " & oDDEConv1.IsOpen()
WScript.Echo "Val : " & oDDEConv1.RequestTextItem(strItem)
WScript.Echo "Conv: " & oDDEConv2.Service & "|" & oDDEConv1.Topic
WScript.Echo "Open? " & oDDEConv2.IsOpen()
WScript.Echo "Conv: " & oDDEConv3.Service & "|" & oDDEConv1.Topic
WScript.Echo "Open? " & oDDEConv3.IsOpen()

WScript.Echo ""
WScript.Echo "Get the collection of conversations"
WScript.Echo "------------------------------------------------------------"

' Query for the collection of open conversations.
Dim oDDEConvSet
Set oDDEConvSet = oDDEClient.Conversations()

WScript.Echo "Count: " & oDDEConvSet.Count

WScript.Echo ""
WScript.Echo "Iterate the collection of conversations by indexing"
WScript.Echo "------------------------------------------------------------"

Dim i, j

' Iterate by indexing.
For i = 0 To oDDEConvSet.Count-1

	Set oDDEConv = oDDEConvSet.Item(i)
	WScript.Echo "Conv: " & oDDEConv.Service & "|" & oDDEConv.Topic

Next

WScript.Echo ""
WScript.Echo "Iterate the collection of conversations by enumeration"
WScript.Echo "------------------------------------------------------------"

' Iterate by using an enumerator.
For Each oDDEConv In oDDEConvSet

	WScript.Echo "Conv: " & oDDEConv.Service & "|" & oDDEConv.Topic

Next

WScript.Echo ""
WScript.Echo "Open a conversation to an invalid service"
WScript.Echo "------------------------------------------------------------"

On Error Resume Next

' Create a non-existent conversation to test the error handling.
Set oDDEConv = oDDEClient.OpenConversation("UNKNOWN", "UNKNOWN")

WScript.Echo Err.Source & " : " & Err.Description

On Error Goto 0

WScript.Echo ""
WScript.Echo "Create a conversation using a 'ddelink:' moniker"
WScript.Echo "------------------------------------------------------------"

' Create a conversation from a moniker.
Set oDDEConv = GetObject("ddelink://PROGMAN|PROGMAN")

WScript.Echo "Conv: " & oDDEConv.Service & "|" & oDDEConv.Topic
WScript.Echo "Val : " & oDDEConv.RequestTextItem(strItem)

WScript.Echo ""
WScript.Echo "Enumerate the collection of running servers and topics"
WScript.Echo "------------------------------------------------------------"

Dim astrServers, astrTopics

' Get the list of running DDE servers and their topics.
astrServers = oDDEClient.RunningServers()

WScript.Echo "Count: " & UBound(astrServers) - LBound(astrServers) + 1
WScript.Echo ""

For i = LBound(astrServers) To UBound(astrServers)

	WScript.Echo "Server: " & astrServers(i)

	astrTopics = oDDEClient.GetServerTopics(astrServers(i))

	For j = LBound(astrTopics) To UBound(astrTopics)

		WScript.Echo "    Topic: " & astrTopics(j)

	Next

	WScript.Echo ""

Next

' Cleanup
Set oDDEClient  = nothing
Set oDDEConv    = nothing
Set oDDEConv1   = nothing
Set oDDEConv2   = nothing
Set oDDEConv3   = nothing
Set oDDEConvSet = nothing

WScript.Echo "------------------------------------------------------------"
WScript.Echo "Tests completed"
