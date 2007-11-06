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
	HRESULT hr = S_OK;

	// Do the standard registration.
	hr = InprocServer::DllRegisterServer();

	// Do the custom registration.
	ComDDEClassFactory::RegisterNamespace();

	return hr;
}

////////////////////////////////////////////////////////////////////////////////
//! Unregister the server from the registry.

HRESULT ComServer::DllUnregisterServer()
{
	HRESULT hr = S_OK;

	// Do the custom unregistration.
	ComDDEClassFactory::UnregisterNamespace();

	// Do the standard unregistration.
	hr = InprocServer::DllUnregisterServer();

	return hr;
}

////////////////////////////////////////////////////////////////////////////////
//! Template Method to get the servers class factory.

COM::IClassFactoryPtr ComServer::CreateClassFactory(const CLSID& oCLSID)
{
	return COM::IClassFactoryPtr(new ComDDEClassFactory(oCLSID), true);
}
