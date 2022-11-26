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

Releases
--------

Stable releases are formally packaged and made available from my Win32 tools page:
http://www.chrisoldwood.com/win32.htm

The latest code is available from my GitHub repo:
https://github.com/chrisoldwood/DDECOMClient

Installation
------------

Register the COM component using the version of regsvr32.exe that matches the
relevant Windows architecture.

32-bit DLL on 32-bit Windows:

> regsvr32.exe DDECOMClient32.dll

64-bit DLL on 64-bit Windows:

> regsvr32.exe DDECOMClient64.dll

32-bit DLL on 64-bit Windows:

> %SystemRoot%\SysWOW64\regsvr32.exe DDECOMClient32.dll

Uninstallation
--------------

Unregister the component using the relevant regsvr32.exe.

> regsvr32.exe /u DDECOMClient<32|64>.dll

Documentation
-------------

There is an HTML based manual - DDECOMClient.html.

Contact Details
------------------

Email: gort@cix.co.uk
Web:   http://www.chrisoldwood.com

Chris Oldwood 
8th August 2013
