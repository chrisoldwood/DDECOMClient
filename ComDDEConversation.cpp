////////////////////////////////////////////////////////////////////////////////
//! \file   ComDDEConversation.cpp
//! \brief  The ComDDEConversation class definition.
//! \author Chris Oldwood

#include "DDECOMClient.hpp"
#include "ComDDEConversation.hpp"
#include <NCL/DDEException.hpp>
#include <NCL/DDEClient.hpp>
#include <WCL/VariantBool.hpp>

#ifdef _DEBUG
// For memory leak detection.
#define new DBGCRT_NEW
#endif

////////////////////////////////////////////////////////////////////////////////
//! Default constructor.

ComDDEConversation::ComDDEConversation()
	: COM::IDispatchImpl<ComDDEConversation>(IID_IDDEConversation)
	, m_pConv(nullptr)
{
}

////////////////////////////////////////////////////////////////////////////////
//! Construction from a Service and Topic name.

ComDDEConversation::ComDDEConversation(const std::tstring& strService, const std::tstring& strTopic)
	: COM::IDispatchImpl<ComDDEConversation>(IID_IDDEConversation)
	, m_strService(strService)
	, m_strTopic(strTopic)
	, m_pConv(nullptr)
{
}

////////////////////////////////////////////////////////////////////////////////
//! Construction from a DDE conversation.

ComDDEConversation::ComDDEConversation(DDE::CltConvPtr pConv)
	: COM::IDispatchImpl<ComDDEConversation>(IID_IDDEConversation)
	, m_strService(pConv->Service())
	, m_strTopic(pConv->Topic())
	, m_pConv(pConv)
{
}

////////////////////////////////////////////////////////////////////////////////
//! Destructor.

ComDDEConversation::~ComDDEConversation()
{
}

////////////////////////////////////////////////////////////////////////////////
//! Set the Service name.

HRESULT COMCALL ComDDEConversation::put_Service(BSTR bstrService)
{
	HRESULT hr = E_FAIL;

	try
	{
		USES_CONVERSION;

		// Currently open?
		if (m_pConv.Get() != nullptr)
			throw WCL::ComException(E_UNEXPECTED, "The conversation is open");

		// Validate parameters.
		if (bstrService == nullptr)
			throw WCL::ComException(E_POINTER, "bstrService is NULL");

		// Save service name.
		m_strService = OLE2T(bstrService);

		hr = S_OK;
	}
	COM_CATCH(hr)

	return hr;
}

////////////////////////////////////////////////////////////////////////////////
//! Get the Service name.

HRESULT COMCALL ComDDEConversation::get_Service(BSTR* pbstrService)
{
	HRESULT hr = E_FAIL;

	try
	{
		USES_CONVERSION;

		// Validate parameters.
		if (pbstrService == nullptr)
			throw WCL::ComException(E_POINTER, "pbstrService is NULL");

		// Return service name.
		*pbstrService = ::SysAllocString(T2OLE(m_strService.c_str()));

		hr = S_OK;
	}
	COM_CATCH(hr)

	return hr;
}

////////////////////////////////////////////////////////////////////////////////
//! Set the Topic name.

HRESULT COMCALL ComDDEConversation::put_Topic(BSTR bstrTopic)
{
	HRESULT hr = E_FAIL;

	try
	{
		USES_CONVERSION;

		// Currently open?
		if (m_pConv.Get() != nullptr)
			throw WCL::ComException(E_UNEXPECTED, "The conversation is open");

		// Validate parameters.
		if (bstrTopic == nullptr)
			throw WCL::ComException(E_POINTER, "bstrTopic is NULL");

		// Save topic name.
		m_strTopic = OLE2T(bstrTopic);

		hr = S_OK;
	}
	COM_CATCH(hr)

	return hr;
}

////////////////////////////////////////////////////////////////////////////////
//! Get the Topic name.

HRESULT COMCALL ComDDEConversation::get_Topic(BSTR* pbstrTopic)
{
	HRESULT hr = E_FAIL;

	try
	{
		USES_CONVERSION;

		// Validate parameters.
		if (pbstrTopic == nullptr)
			throw WCL::ComException(E_POINTER, "pbstrTopic is NULL");

		// Return topic name.
		*pbstrTopic = ::SysAllocString(T2OLE(m_strTopic.c_str()));

		hr = S_OK;
	}
	COM_CATCH(hr)

	return hr;
}

////////////////////////////////////////////////////////////////////////////////
//! Open the conversation.

HRESULT COMCALL ComDDEConversation::Open()
{
	HRESULT hr = E_FAIL;

	try
	{
		// Already open?
		if (m_pConv.Get() != nullptr)
			throw WCL::ComException(E_UNEXPECTED, "The conversation is already open");

		// Open it.
		m_pConv = DDE::CltConvPtr(CDDEClient::Instance()->CreateConversation(m_strService.c_str(), m_strTopic.c_str()));

		hr = S_OK;
	}
	COM_CATCH(hr)

	return hr;
}

////////////////////////////////////////////////////////////////////////////////
//! Query if the conversation is open.

HRESULT COMCALL ComDDEConversation::IsOpen(VARIANT_BOOL* pbIsOpen)
{
	HRESULT hr = E_FAIL;

	try
	{
		// Validate parameters.
		if (pbIsOpen == nullptr)
			throw WCL::ComException(E_POINTER, "pbIsOpen is NULL");

		// Return conversation status.
		*pbIsOpen = ToVariantBool(m_pConv.Get() != nullptr);

		hr = S_OK;
	}
	COM_CATCH(hr)

	return hr;
}

////////////////////////////////////////////////////////////////////////////////
//! Close the conversation.

HRESULT COMCALL ComDDEConversation::Close()
{
	HRESULT hr = E_FAIL;

	try
	{
		// Destroy the conversation.
		m_pConv.Release();

		hr = S_OK;
	}
	COM_CATCH(hr)

	return hr;
}

////////////////////////////////////////////////////////////////////////////////
//! Request an item in CF_TEXT format.

HRESULT COMCALL ComDDEConversation::RequestTextItem(BSTR bstrItem, BSTR* pbstrValue)
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
		if (bstrItem == nullptr)
			throw WCL::ComException(E_INVALIDARG, "bstrItem is NULL");

		// Conversation not open?
		if (m_pConv.Get() == nullptr)
			throw WCL::ComException(E_UNEXPECTED, "The conversation is not open");

		// Request the item value (CF_ANSI).
		std::tstring strItem  = OLE2T(bstrItem);
		std::tstring strValue = m_pConv->Request(strItem.c_str());

		// Return value.
		*pbstrValue = ::SysAllocString(T2OLE(strValue.c_str()));

		hr = S_OK;
	}
	COM_CATCH(hr)

	return hr;
}
