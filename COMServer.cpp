////////////////////////////////////////////////////////////////////////////////
//! \file   ComServer.cpp
//! \brief  The ComServer class definition.
//! \author Chris Oldwood

#include "Common.hpp"
#include "ComServer.hpp"

//! The component object.
ComServer g_oDll;

////////////////////////////////////////////////////////////////////////////////
//! Default constructor.

ComServer::ComServer()
{
}

////////////////////////////////////////////////////////////////////////////////
//! Destructor.

ComServer::~ComServer()
{
}

////////////////////////////////////////////////////////////////////////////////
//! Register the server in the registry.

HRESULT ComServer::DllRegisterServer()
{
	HRESULT hr = E_FAIL;

	// Do the standard registration.
	hr = InprocServer::DllRegisterServer();

	if (FAILED(hr))
		return hr;

	// Do the custom registration.
	ComDDEClassFactory::RegisterNamespace(COM::MACHINE);

	return hr;
}

////////////////////////////////////////////////////////////////////////////////
//! Unregister the server from the registry.

HRESULT ComServer::DllUnregisterServer()
{
	HRESULT hr = E_FAIL;

	// Do the custom unregistration.
	ComDDEClassFactory::UnregisterNamespace(COM::MACHINE);

	// Do the standard unregistration.
	hr = InprocServer::DllUnregisterServer();

	return hr;
}

////////////////////////////////////////////////////////////////////////////////
//! Register or unregister the server to/from the registry.

HRESULT ComServer::DllInstall(bool install, const tchar* cmdLine)
{
	HRESULT hr = E_FAIL;

	hr = InprocServer::DllInstall(install, cmdLine);

	if (FAILED(hr))
		return hr;

	bool perUser = ( (cmdLine != nullptr) && (tstricmp(cmdLine, TXT("user")) == 0) );
	COM::Scope scope = (perUser) ? COM::USER : COM::MACHINE;

	(install) ? ComDDEClassFactory::RegisterNamespace(scope)
	          : ComDDEClassFactory::UnregisterNamespace(scope);

	return hr;
}

////////////////////////////////////////////////////////////////////////////////
//! Template Method to get the servers class factory.

COM::IClassFactoryPtr ComServer::CreateClassFactory(const CLSID& oCLSID)
{
	return COM::IClassFactoryPtr(new ComDDEClassFactory(oCLSID), true);
}
