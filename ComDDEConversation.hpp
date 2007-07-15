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
	ComDDEConversation(const std::tstring& strService, const std::tstring& strTopic);

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

	//! Open the conversation.
	virtual HRESULT COMCALL Open();

	//! Query if the conversation is open.
	virtual HRESULT COMCALL IsOpen(VARIANT_BOOL* pbIsOpen);

	//! Close the conversation.
	virtual HRESULT COMCALL Close();

	//! Request an item in CF_TEXT format.
	virtual HRESULT COMCALL RequestTextItem(BSTR bstrItem, BSTR* pbstrValue);

private:
	//
	// Members.
	//
	std::tstring	m_strService;		//!< The service name.
	std::tstring	m_strTopic;			//!< The topic name.
	DDE::CltConvPtr	m_pConv;			//!< The underlying DDE conversation.
};

#endif // COMDDECONVERSATION_HPP
