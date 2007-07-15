////////////////////////////////////////////////////////////////////////////////
//! \file   ComDDEClient.cpp
//! \brief  The ComDDEClient class definition.
//! \author Chris Oldwood

#include "DDECOMClient.hpp"
#include "ComDDEClient.hpp"
#include <NCL/DDECltConv.hpp>
#include <NCL/DDECltConvPtr.hpp>
#include <NCL/DDEException.hpp>
#include "ComDDEConversation.hpp"
#include "ComDDEConversations.hpp"
#include <WCL/SafeVector.hpp>
#include <WCL/StrArray.hpp>

#ifdef _DEBUG
// For memory leak detection.
#define new DBGCRT_NEW
#endif

////////////////////////////////////////////////////////////////////////////////
//! Default constructor.

ComDDEClient::ComDDEClient()
	: COM::IDispatchImpl<ComDDEClient>(IID_IDDEClient)
	, m_pDDEClient(ComDDEClient::DDEClient())
{
}

////////////////////////////////////////////////////////////////////////////////
//! Destructor.

ComDDEClient::~ComDDEClient()
{
}

////////////////////////////////////////////////////////////////////////////////
//! Query for the collection of all running servers.

HRESULT COMCALL ComDDEClient::RunningServers(SAFEARRAY** ppServers)
{
	HRESULT hr = E_FAIL;

	try
	{
		// Type shorthands.
		typedef WCL::SafeVector<VARIANT> VariantArray;

		USES_CONVERSION;

		// Check output parameters.
		if (ppServers == nullptr)
			throw WCL::ComException(E_POINTER, "ppServers is NULL");

		// Reset output parameters.
		*ppServers = nullptr;

		CStrArray astrServers;

		// Get the list of running servers.
		m_pDDEClient->QueryServers(astrServers);

		VariantArray           avtServers(astrServers.Size());
		VariantArray::iterator itServer = avtServers.begin();

		// Create the return array.
		for (int i = 0; i < astrServers.Size(); ++i, ++itServer)
		{
			ASSERT(itServer != avtServers.end());

			::VariantInit(itServer);

			V_VT  (itServer) = VT_BSTR;
			V_BSTR(itServer) = ::SysAllocString(T2OLE(astrServers[i]));
		}

		// Return value.
		*ppServers = avtServers.Detach();

		hr = S_OK;
	}
	COM_CATCH(hr)

	return hr;
}

////////////////////////////////////////////////////////////////////////////////
//! Query for the topics supported by the server.

HRESULT COMCALL ComDDEClient::GetServerTopics(BSTR bstrService, SAFEARRAY** ppTopics)
{
	HRESULT hr = E_FAIL;

	try
	{
		// Type shorthands.
		typedef WCL::SafeVector<VARIANT> VariantArray;

		USES_CONVERSION;

		// Check output parameters.
		if (ppTopics == nullptr)
			throw WCL::ComException(E_POINTER, "ppTopics is NULL");

		// Reset output parameters.
		*ppTopics = nullptr;

		// Validate input parameters.
		if (bstrService == nullptr)
			throw WCL::ComException(E_INVALIDARG, "bstrService is NULL");

		std::tstring strService = OLE2T(bstrService);

		CStrArray astrTopics;

		// Get the list of topics for the server.
		m_pDDEClient->QueryServerTopics(strService.c_str(), astrTopics);

		VariantArray           avtTopics(astrTopics.Size());
		VariantArray::iterator itTopic = avtTopics.begin();

		// Create the return array.
		for (int i = 0; i < astrTopics.Size(); ++i, ++itTopic)
		{
			ASSERT(itTopic != avtTopics.end());

			::VariantInit(itTopic);

			V_VT  (itTopic) = VT_BSTR;
			V_BSTR(itTopic) = ::SysAllocString(T2OLE(astrTopics[i]));
		}

		// Return value.
		*ppTopics = avtTopics.Detach();

		hr = S_OK;
	}
	COM_CATCH(hr)

	return hr;
}

////////////////////////////////////////////////////////////////////////////////
//! Open a conversation.

HRESULT COMCALL ComDDEClient::OpenConversation(BSTR bstrService, BSTR bstrTopic, IDDEConversation** ppIDDEConv)
{
	HRESULT hr = E_FAIL;

	try
	{
		// Type shorthands.
		typedef Core::IFacePtr<IDDEConversation> IDDEConversationPtr;

		USES_CONVERSION;

		// Check output parameters.
		if (ppIDDEConv == nullptr)
			throw WCL::ComException(E_POINTER, "ppIDDEConv is NULL");

		// Reset output parameters.
		*ppIDDEConv = nullptr;

		// Validate input parameters.
		if ( (bstrService == nullptr) || (bstrTopic == nullptr) )
			throw WCL::ComException(E_INVALIDARG, "bstrService/bstrTopic is NULL");

		// Create the conversation.
		std::tstring strService = OLE2T(bstrService);
		std::tstring strTopic   = OLE2T(bstrTopic);

		DDE::CltConvPtr     pConv    = DDE::CltConvPtr(m_pDDEClient->CreateConversation(strService.c_str(), strTopic.c_str()));
		IDDEConversationPtr pComConv = IDDEConversationPtr(new ComDDEConversation(pConv), true);

		// Return value.
		*ppIDDEConv = pComConv.Detach();

		hr = S_OK;
	}
	COM_CATCH(hr)

	return hr;
}

////////////////////////////////////////////////////////////////////////////////
//! Query for the collection of all open conversations.

HRESULT COMCALL ComDDEClient::Conversations(IDDEConversations** ppIDDEConvs)
{
	HRESULT hr = E_FAIL;

	try
	{
		// Type shorthands.
		typedef Core::IFacePtr<IDDEConversations> IDDEConversationsPtr;

		// Check output parameters.
		if (ppIDDEConvs == nullptr)
			throw WCL::ComException(E_POINTER, "ppIDDEConvs is NULL");

		// Reset output parameters.
		*ppIDDEConvs = nullptr;

		CDDECltConvs apConvs;

		// Get all open conversations.
		CDDEClient::Instance()->GetAllConversations(apConvs);

		DDECltConvs vpCltConvs(apConvs.size());

		// Convert to a vector of smart-pointers.
		// NB: Need to AddRef the returned pointers.
		for (size_t i = 0; i < apConvs.size(); ++i)
			vpCltConvs[i] = DDE::CltConvPtr(apConvs[i], true);

		IDDEConversationsPtr pComConvs = IDDEConversationsPtr(new ComDDEConversations(vpCltConvs), true);

		// Return value.
		*ppIDDEConvs = pComConvs.Detach();

		hr = S_OK;
	}
	COM_CATCH(hr)

	return hr;
}

////////////////////////////////////////////////////////////////////////////////
//! Request an item in CF_TEXT format.

HRESULT COMCALL ComDDEClient::RequestTextItem(BSTR bstrService, BSTR bstrTopic, BSTR bstrItem, BSTR* pbstrValue)
{
	HRESULT hr = E_FAIL;

	try
	{
		USES_CONVERSION;

		// Check output parameters.
		if (pbstrValue == nullptr)
			throw WCL::ComException(E_POINTER, "pbstrValue is NULL");

		// Reset output parameters.
		*pbstrValue = nullptr;

		// Validate input parameters.
		if ( (bstrService == nullptr) || (bstrTopic == nullptr) || (bstrItem == nullptr) )
			throw WCL::ComException(E_INVALIDARG, "bstrService/bstrTopic/bstrItem is NULL");

		// Create the conversation.
		std::tstring strService = OLE2T(bstrService);
		std::tstring strTopic   = OLE2T(bstrTopic);

		DDE::CltConvPtr pConv(m_pDDEClient->CreateConversation(strService.c_str(), strTopic.c_str()));

		// Request the item value (CF_ANSI).
		std::tstring strItem  = OLE2T(bstrItem);
		std::tstring strValue = pConv->Request(strItem.c_str());

		// Return value.
		*pbstrValue = ::SysAllocString(T2OLE(strValue.c_str()));

		hr = S_OK;
	}
	COM_CATCH(hr)

	return hr;
}

////////////////////////////////////////////////////////////////////////////////
//! Aquire the singleton DDE Client.

CDDEClient* ComDDEClient::DDEClient()
{
	CDDEClient* pDDEClient = CDDEClient::Instance();

	// Initialise on first use.
	if (pDDEClient->Handle() == 0)
		pDDEClient->Initialise();

	return pDDEClient;
}
