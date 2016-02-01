////////////////////////////////////////////////////////////////////////////////
//! \file   ComDDEClassFactory.hpp
//! \brief  The ComDDEClassFactory class declaration.
//! \author Chris Oldwood

// Check for previous inclusion
#ifndef COMDDECLASSFACTORY_HPP
#define COMDDECLASSFACTORY_HPP

#include <COM/ClassFactory.hpp>
#include <COM/RegUtils.hpp>

#if _MSC_VER > 1000
#pragma once
#endif

////////////////////////////////////////////////////////////////////////////////
//! The custom Class Factory which can also create objects from item monikers.

class ComDDEClassFactory : public COM::ClassFactory, public IOleItemContainer

{
public:
	//! Default constructor.
	ComDDEClassFactory();

	//! Construction from a CLSID.
	ComDDEClassFactory(const CLSID& rCLSID);

	//! Destructor.
	virtual ~ComDDEClassFactory();
	
	DEFINE_INTERFACE_TABLE(IClassFactory)
		IMPLEMENT_INTERFACE(IID_IClassFactory,     IClassFactory)
		IMPLEMENT_INTERFACE(IID_IParseDisplayName, IParseDisplayName)
		IMPLEMENT_INTERFACE(IID_IOleContainer,     IOleContainer)
		IMPLEMENT_INTERFACE(IID_IOleItemContainer, IOleItemContainer)
	END_INTERFACE_TABLE()

	IMPLEMENT_IUNKNOWN()

	//
	// IParseDisplayName methods.
	//

	//! Parse the string and create a moniker for it.
	virtual HRESULT COMCALL ParseDisplayName(IBindCtx* pBindCtx, LPOLESTR pszDisplayName, ULONG* pcEaten, IMoniker** ppMoniker);

	//
	// IOleContainer methods.
	//

	//! Enumerate the objects in the container.
	virtual HRESULT COMCALL EnumObjects(DWORD dwFlags, IEnumUnknown** ppEnum);

	//! Lock the container to ensure it is kept active.
	virtual HRESULT COMCALL LockContainer(BOOL bLock);

	//
	// IOleItemContainer methods.
	//

	//! Get the object associated with the item.
	virtual HRESULT COMCALL GetObject(LPOLESTR pszItem, DWORD dwSpeedNeeded, IBindCtx* pBindCtx, REFIID rIID, void** ppObject);

	//! Get the storage associated with the item.
	virtual HRESULT COMCALL GetObjectStorage(LPOLESTR pszItem, IBindCtx* pBindCtx, REFIID rIID, void** ppStorage);

	//! Query if the object is already running.
	virtual HRESULT COMCALL IsRunning(LPOLESTR pszItem);

	//
	// Registration methods.
	//

	//! Register the moniker namespace in the registry.
	static void RegisterNamespace(COM::Scope scope);

	//! Unregister the moniker namespace from the registry.
	static void UnregisterNamespace(COM::Scope scope);

private:
	//
	// Members.
	//
};

#endif // COMDDECLASSFACTORY_HPP
