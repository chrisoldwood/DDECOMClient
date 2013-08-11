////////////////////////////////////////////////////////////////////////////////
//! \file   ComDDEConversation.hpp
//! \brief  The ComDDEConversation class declaration.
//! \author Chris Oldwood

// Check for previous inclusion
#ifndef COMDDECONVERSATION_HPP
#define COMDDECONVERSATION_HPP

#include <COM/IDispatchImpl.hpp>
#include <NCL/DDECltConv.hpp>
#include <NCL/DDECltConvPtr.hpp>

#if _MSC_VER > 1000
#pragma once
#endif

////////////////////////////////////////////////////////////////////////////////
//! The COM facade over the underlying NCL DDEConv object.

class ComDDEConversation :	public  COM::ObjectBase<IDDEConversation>,
							private COM::IDispatchImpl<ComDDEConversation>
{
public:
	//! Default constructor.
	ComDDEConversation();

	//! Construction from a Service and Topic name.
	ComDDEConversation(const tstring& strService, const tstring& strTopic);

	//! Construction from a DDE conversation.
	ComDDEConversation(DDE::CltConvPtr pConv);

	//! Destructor.
	virtual ~ComDDEConversation();
	
	DEFINE_INTERFACE_TABLE(IDDEConversation)
		IMPLEMENT_INTERFACE(IID_IDDEConversation,  IDDEConversation)
		IMPLEMENT_INTERFACE(IID_IDispatch,         IDDEConversation)
	END_INTERFACE_TABLE()

	IMPLEMENT_IUNKNOWN()
	IMPLEMENT_IDISPATCH(ComDDEConversation)

	//
	// IDDEConversation methods.
	//

	//! Set the Service name.
	virtual HRESULT COMCALL put_Service(BSTR bstrService);

	//! Get the Service name.
	virtual HRESULT COMCALL get_Service(BSTR* pbstrService);

	//! Set the Topic name.
	virtual HRESULT COMCALL put_Topic(BSTR bstrTopic);

	//! Get the Topic name.
	virtual HRESULT COMCALL get_Topic(BSTR* pbstrTopic);

	//! Set the maximum time (ms) to wait for a reply.
	virtual HRESULT COMCALL put_Timeout(long timeout);

	//! Get the maximum time (ms) to wait for a reply.
	virtual HRESULT COMCALL get_Timeout(long* timeout);

	//! Open the conversation.
	virtual HRESULT COMCALL Open();

	//! Query if the conversation is open.
	virtual HRESULT COMCALL IsOpen(VARIANT_BOOL* pbIsOpen);

	//! Close the conversation.
	virtual HRESULT COMCALL Close();

	//! Request an item in CF_TEXT format.
	virtual HRESULT COMCALL RequestTextItem(BSTR bstrItem, BSTR* pbstrValue);

	//! Poke an item provided in CF_TEXT format.
	virtual HRESULT COMCALL PokeTextItem(BSTR bstrItem, BSTR bstrValue);

	//! Execute a command provided in CF_TEXT/CF_UNICODETEXT format.
	virtual HRESULT COMCALL ExecuteTextCommand(BSTR bstrCommand);

private:
	//
	// Members.
	//
	tstring			m_strService;		//!< The service name.
	tstring			m_strTopic;			//!< The topic name.
	DWORD			m_timeout;			//!< The conversation timeout.
	DDE::ClientPtr	m_pDDEClient;		//!< The DDE client.
	DDE::CltConvPtr	m_pConv;			//!< The underlying DDE conversation.
};

#endif // COMDDECONVERSATION_HPP
