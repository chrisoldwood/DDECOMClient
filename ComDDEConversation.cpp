////////////////////////////////////////////////////////////////////////////////
//! \file   ComDDEConversation.cpp
//! \brief  The ComDDEConversation class definition.
//! \author Chris Oldwood

#include "Common.hpp"
#include "ComDDEConversation.hpp"
#include <NCL/DDEException.hpp>
#include <NCL/DDEClient.hpp>
#include <WCL/VariantBool.hpp>

////////////////////////////////////////////////////////////////////////////////
//! Default constructor.

ComDDEConversation::ComDDEConversation()
	: COM::IDispatchImpl<ComDDEConversation>(IID_IDDEConversation)
	, m_pConv(nullptr)
{
}

////////////////////////////////////////////////////////////////////////////////
//! Construction from a Service and Topic name.

ComDDEConversation::ComDDEConversation(const tstring& strService, const tstring& strTopic)
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
		// Currently open?
		if (m_pConv.Get() != nullptr)
			throw WCL::ComException(E_UNEXPECTED, TXT("The conversation is open"));

		// Validate parameters.
		if (bstrService == nullptr)
			throw WCL::ComException(E_POINTER, TXT("bstrService is NULL"));

		// Save service name.
		m_strService = W2T(bstrService);

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
		// Validate parameters.
		if (pbstrService == nullptr)
			throw WCL::ComException(E_POINTER, TXT("pbstrService is NULL"));

		// Return service name.
		*pbstrService = ::SysAllocString(T2W(m_strService.c_str()));

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
		// Currently open?
		if (m_pConv.Get() != nullptr)
			throw WCL::ComException(E_UNEXPECTED, TXT("The conversation is open"));

		// Validate parameters.
		if (bstrTopic == nullptr)
			throw WCL::ComException(E_POINTER, TXT("bstrTopic is NULL"));

		// Save topic name.
		m_strTopic = W2T(bstrTopic);

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
		// Validate parameters.
		if (pbstrTopic == nullptr)
			throw WCL::ComException(E_POINTER, TXT("pbstrTopic is NULL"));

		// Return topic name.
		*pbstrTopic = ::SysAllocString(T2W(m_strTopic.c_str()));

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
			throw WCL::ComException(E_UNEXPECTED, TXT("The conversation is already open"));

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
			throw WCL::ComException(E_POINTER, TXT("pbIsOpen is NULL"));

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
		// Check output parameters.
		if (pbstrValue == nullptr)
			throw WCL::ComException(E_POINTER, TXT("pbstrValue is NULL"));

		// Reset output parameters.
		*pbstrValue = nullptr;

		// Validate input parameters.
		if (bstrItem == nullptr)
			throw WCL::ComException(E_INVALIDARG, TXT("bstrItem is NULL"));

		// Conversation not open?
		if (m_pConv.Get() == nullptr)
			throw WCL::ComException(E_UNEXPECTED, TXT("The conversation is not open"));

		// Request the item value (CF_ANSI).
		tstring strItem  = W2T(bstrItem);
		tstring strValue = m_pConv->RequestString(strItem.c_str(), CF_TEXT);

		// Return value.
		*pbstrValue = ::SysAllocString(T2W(strValue.c_str()));

		hr = S_OK;
	}
	COM_CATCH(hr)

	return hr;
}
