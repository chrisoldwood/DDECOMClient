////////////////////////////////////////////////////////////////////////////////
//! \file   ComServer.hpp
//! \brief  The ComServer class declaration.
//! \author Chris Oldwood

// Check for previous inclusion
#ifndef COMSERVER_HPP
#define COMSERVER_HPP

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

	DEFINE_REGISTRATION_TABLE("DDECOMClient", LIBID_DDECOMClientLib, 1, 0)
		DEFINE_CLASS_REG_INFO(CLSID_DDEClient,       "DDEClient",       "1", COM::MAIN_THREAD_APT)
		DEFINE_CLASS_REG_INFO(CLSID_DDEConversation, "DDEConversation", "1", COM::MAIN_THREAD_APT)
		DEFINE_CLASS_REG_INFO(CLSID_DDEClassFactory, "DDEClassFactory", "1", COM::MAIN_THREAD_APT)
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
