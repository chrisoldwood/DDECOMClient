////////////////////////////////////////////////////////////////////////////////
//! \file   ComServer.hpp
//! \brief  The ComServer class declaration.
//! \author Chris Oldwood

// Check for previous inclusion
#ifndef COMSERVER_HPP
#define COMSERVER_HPP

#if _MSC_VER > 1000
#pragma once
#endif

#include <COM/InprocServer.hpp>
#include <COM/ServerRegInfo.hpp>
#include "TypeLibrary_h.h"
#include "ComDDEClient.hpp"
#include "ComDDEConversation.hpp"
#include "ComDDEClassFactory.hpp"

////////////////////////////////////////////////////////////////////////////////
//! The COM server concrete class.

class ComServer : public COM::InprocServer
{
public:
	//! Default constructor.
	ComServer();

	//! Destructor.
	virtual ~ComServer();

protected:
	//	
	// Overriden COM::InprocServer methods.
	//

	//! Register the server in the registry.
	virtual HRESULT DllRegisterServer();

	//! Unregister the server from the registry.
	virtual HRESULT DllUnregisterServer();

	//! Register or unregister the server to/from the registry.
	virtual HRESULT DllInstall(bool install, const tchar* cmdLine);

	DEFINE_REGISTRATION_TABLE(TXT("DDECOMClient"), LIBID_DDECOMClientLib, 1, 0)
		DEFINE_CLASS_REG_INFO(CLSID_DDEClient,       TXT("DDEClient"),       TXT("1"), COM::MAIN_THREAD_APT)
		DEFINE_CLASS_REG_INFO(CLSID_DDEConversation, TXT("DDEConversation"), TXT("1"), COM::MAIN_THREAD_APT)
		DEFINE_CLASS_REG_INFO(CLSID_DDEClassFactory, TXT("DDEClassFactory"), TXT("1"), COM::MAIN_THREAD_APT)
	END_REGISTRATION_TABLE()

	DEFINE_CLASS_FACTORY_TABLE()
		DEFINE_CLASS(CLSID_DDEClient,       ComDDEClient,       IDDEClient)
		DEFINE_CLASS(CLSID_DDEConversation, ComDDEConversation, IDDEConversation)
		DEFINE_CLASS(CLSID_DDEClassFactory, ComDDEClassFactory, IClassFactory)
	END_CLASS_FACTORY_TABLE()

	//! Template Method to get the servers class factory.
	virtual COM::IClassFactoryPtr CreateClassFactory(const CLSID& oCLSID);
};

////////////////////////////////////////////////////////////////////////////////
// Global variables.

//! The component object.
extern ComServer g_oDll;

#endif // COMSERVER_HPP
