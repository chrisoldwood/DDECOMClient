DDE COM Client v1.1.1
=====================

Introduction
------------

This is a COM component designed to allow access to DDE servers from Scripting
languages, such as VBScript. It provides support for the main DDE synchronous
transaction types: XTYP_REQUEST, XTYP_POKE and XTYP_EXECUTE through ad-hoc
queries and conversations. It also allows for querying running servers with
the XTYP_WILDCONNECT transaction type.

This tool forms one part of my DDE toolkit, with the others being:-

DDE Query: My original GUI based DDE query tool.
DDE Command: A console app for querying DDE servers.

Installation
------------

Register the component using "regsvr32.exe DDECOMClient.dll"

Uninstallation
--------------

Unregister the component using "regsvr32.exe /u DDECOMClient.dll"

Documentation
-------------

There is an HTML based manual - DDECOMClient.html.

Contact Details
------------------

Email: gort@cix.co.uk
Web:   http://www.chrisoldwood.com

Chris Oldwood 
8th August 2013
