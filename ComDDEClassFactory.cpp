////////////////////////////////////////////////////////////////////////////////
//! \file   ComDDEClassFactory.cpp
//! \brief  The ComDDEClassFactory class definition.
//! \author Chris Oldwood

#include "Common.hpp"
#include "ComDDEClassFactory.hpp"
#include "TypeLibrary_h.h"
#include <COM/RegUtils.hpp>
#include <WCL/StrTok.hpp>
#include "ComDDEClient.hpp"
#include <NCL/DDECltConvPtr.hpp>
#include "ComDDEConversation.hpp"

////////////////////////////////////////////////////////////////////////////////
// Contants.

//! The moniker namespace registry key.
static const tchar* g_pszRegKey = TXT("DDELINK");

//! The moniker namespace.
static const wchar_t* g_pszNamespace = L"ddelink:";

//! The length of the namespace string.
static const size_t g_nNamespaceLen = wcslen(g_pszNamespace);

//! The namespace/item separator string.
static const wchar_t* g_pszDelim = L"//";

//! The length of the separator string.
static const size_t g_nDelimLen = wcslen(g_pszDelim);

////////////////////////////////////////////////////////////////////////////////
//! Default constructor. This needs to be created by the server class factory
//! without a CLSID to obtain the IParseDisplayName interface.

ComDDEClassFactory::ComDDEClassFactory()
	: COM::ClassFactory(CLSID_NULL)
{
}

////////////////////////////////////////////////////////////////////////////////
//! Construction from a CLSID.

ComDDEClassFactory::ComDDEClassFactory(const CLSID& rCLSID)
	: COM::ClassFactory(rCLSID)
{
}

////////////////////////////////////////////////////////////////////////////////
//! Destructor.

ComDDEClassFactory::~ComDDEClassFactory()
{
}

////////////////////////////////////////////////////////////////////////////////
//! Parse the string and create a moniker for it. The returned moniker is a
//! composite of a class moniker for the DDE Conversation and an item moniker
//! containing the "service|topic" name.

HRESULT COMCALL ComDDEClassFactory::ParseDisplayName(IBindCtx* pBindCtx, LPOLESTR pszDisplayName, ULONG* pcEaten, IMoniker** ppMoniker)
{
	HRESULT hr = E_FAIL;

	try
	{
		// Type shorthands.
		typedef Core::IFacePtr<IMoniker> IMonikerPtr;

		// Check output parameters.
		if ( (pcEaten == nullptr) || (ppMoniker == nullptr) )
			throw WCL::ComException(E_POINTER, TXT("pcEaten/ppMoniker is NULL"));

		// Reset output parameters.
		*pcEaten   = 0;
		*ppMoniker = nullptr;

		// Validate input parameters.
		if ( (pBindCtx == nullptr) || (pszDisplayName == nullptr) )
			throw WCL::ComException(E_INVALIDARG, TXT("pBindCtx/pszDisplayName is NULL"));

//		TRACE1(TXT("%ls\n"), pszDisplayName);

		size_t nLength = wcslen(pszDisplayName);

		// Namespace must be "ddeconv:"
		if (_wcsnicmp(pszDisplayName, g_pszNamespace, g_nNamespaceLen) != 0)
			throw WCL::ComException(MK_E_NOOBJECT, TXT("Namespace must be 'ddeconv:'"));

		// Delim must be "//"
		if (_wcsnicmp(pszDisplayName+g_nNamespaceLen, g_pszDelim, g_nDelimLen) != 0)
			throw WCL::ComException(MK_E_NOOBJECT, TXT("Namespace separator must be '//'"));

		IMonikerPtr pClassMoniker;

		// Create the class moniker for the DDEConversation.
		hr = ::CreateClassMoniker(CLSID_DDEConversation, AttachTo(pClassMoniker));

		if (FAILED(hr))
			throw WCL::ComException(hr, TXT("Failed to create the DDEConversation class moniker"));

		// Skip to the conversation name.
		LPOLESTR pszItemName = pszDisplayName + g_nNamespaceLen + g_nDelimLen;

		IMonikerPtr pItemMoniker;

		// Create the class moniker for the DDEConversation.
		hr = ::CreateItemMoniker(g_pszDelim, pszItemName, AttachTo(pItemMoniker));

		if (FAILED(hr))
			throw WCL::ComException(hr, TXT("Failed to create the DDE conversation item moniker"));

		IMonikerPtr pCompositeMoniker;

		// Create the composite moniker.
		hr = ::CreateGenericComposite(pClassMoniker.Get(), pItemMoniker.Get(), AttachTo(pCompositeMoniker));

		if (FAILED(hr))
			throw WCL::ComException(hr, TXT("Failed to create the DDE conversation composite moniker"));

		// Write output values.
		*pcEaten   = nLength;
		*ppMoniker = pCompositeMoniker.Detach();

		hr = S_OK;
	}
	COM_CATCH(hr)

	return hr;
}

////////////////////////////////////////////////////////////////////////////////
//! Enumerate the objects in the container.

HRESULT COMCALL ComDDEClassFactory::EnumObjects(DWORD /*dwFlags*/, IEnumUnknown** /*ppEnum*/)
{
	ASSERT_FALSE();
	return E_NOTIMPL;
}

////////////////////////////////////////////////////////////////////////////////
//! Lock the container to ensure it is kept active.

HRESULT COMCALL ComDDEClassFactory::LockContainer(BOOL bLock)
{
	if (bLock)
		AddRefImpl();
	else
		ReleaseImpl();

	return S_OK;
}

////////////////////////////////////////////////////////////////////////////////
//! Get the object associated with the item. This creates a DDE Conversation
//! based on the Service and Topic in the Item Moniker.

HRESULT COMCALL ComDDEClassFactory::GetObject(LPOLESTR pszItem, DWORD /*dwSpeedNeeded*/,
							IBindCtx* /*pBindCtx*/, REFIID rIID, void** ppObject)
{
	HRESULT hr = E_FAIL;

	try
	{
		// Type shorthands.
		typedef Core::IFacePtr<IDDEConversation> IDDEConversationPtr;

		// Check output parameters.
		if (ppObject == nullptr)
			throw WCL::ComException(E_POINTER, TXT("ppObject is NULL"));

		// Reset output parameters.
		*ppObject = nullptr;

		// Validate input parameters.
		if (pszItem == nullptr)
			throw WCL::ComException(E_INVALIDARG, TXT("pszItem is NULL"));

		std::tstring strService;
		std::tstring strTopic;
		std::tstring strItem;

		// Split the item into "SERVICE|TOPIC!ITEM".
		std::tstring strLink = W2T(pszItem);
		CStrTok      oStrTok(strLink.c_str(), TXT("|!"));

		// Extract the Service name.
		if (oStrTok.MoreTokens())
			strService = oStrTok.NextToken();

		// Extract the Topic name.
		if (oStrTok.MoreTokens())
			strTopic = oStrTok.NextToken();

		// Extract the Item name.
		if (oStrTok.MoreTokens())
			strItem = oStrTok.NextToken();

		// Must have at least a Service and Topic.
		if (strService.empty() || strTopic.empty())
			throw WCL::ComException(MK_E_SYNTAX, TXT("The DDE link syntax is invalid"));

		// Create a Conversation.
		CDDEClient*         pDDEClient = ComDDEClient::DDEClient();
		DDE::CltConvPtr     pDDEConv   = DDE::CltConvPtr(pDDEClient->CreateConversation(strService.c_str(), strTopic.c_str()));
		IDDEConversationPtr pComConv   = IDDEConversationPtr(new ComDDEConversation(pDDEConv), true);

		// Return value.
		hr = pComConv->QueryInterface(rIID, ppObject);
	}
	COM_CATCH(hr)

	return hr;
}

////////////////////////////////////////////////////////////////////////////////
//! Get the storage associated with the item.

HRESULT COMCALL ComDDEClassFactory::GetObjectStorage(LPOLESTR /*pszItem*/, IBindCtx* /*pBindCtx*/,
														REFIID /*rIID*/, void** /*ppStorage*/)
{
	ASSERT_FALSE();
	return E_NOTIMPL;
}

////////////////////////////////////////////////////////////////////////////////
//! Query if the object is already running.

HRESULT COMCALL ComDDEClassFactory::IsRunning(LPOLESTR /*pszItem*/)
{
	ASSERT_FALSE();
	return E_NOTIMPL;
}

////////////////////////////////////////////////////////////////////////////////
//! Register the moniker namespace in the registry.

void ComDDEClassFactory::RegisterNamespace()
{
	COM::RegisterMonikerPrefix(g_pszRegKey, TXT("DDEClassFactory"), CLSID_DDEClassFactory);
}

////////////////////////////////////////////////////////////////////////////////
//! Unregister the moniker namespace from the registry.

void ComDDEClassFactory::UnregisterNamespace()
{
	COM::UnregisterMonikerPrefix(g_pszRegKey);
}
