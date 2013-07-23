////////////////////////////////////////////////////////////////////////////////
//! \file   ComDDEClient.hpp
//! \brief  The ComDDEClient class declaration.
//! \author Chris Oldwood

// Check for previous inclusion
#ifndef COMDDECLIENT_HPP
#define COMDDECLIENT_HPP

#include <COM/IDispatchImpl.hpp>
#include <NCL/DDEClient.hpp>

#if _MSC_VER > 1000
#pragma once
#endif

////////////////////////////////////////////////////////////////////////////////
//! The COM facade over the underlying NCL DDEClient object.

class ComDDEClient : public COM::ObjectBase<IDDEClient>, COM::IDispatchImpl<ComDDEClient>
{
public:
	//! Default constructor.
	ComDDEClient();

	//! Destructor.
	virtual ~ComDDEClient();
	
	DEFINE_INTERFACE_TABLE(IDDEClient)
		IMPLEMENT_INTERFACE(IID_IDDEClient, IDDEClient)
		IMPLEMENT_INTERFACE(IID_IDispatch,  IDDEClient)
	END_INTERFACE_TABLE()

	IMPLEMENT_IUNKNOWN()
	IMPLEMENT_IDISPATCH(ComDDEClient)

	//
	// IDDEClient methods.
	//

	//! Set the default maximum time (ms) to wait for a reply.
	virtual HRESULT COMCALL put_DefaultTimeout(long timeout);

	//! Get the default maximum time (ms) to wait for a reply.
	virtual HRESULT COMCALL get_DefaultTimeout(long* timeout);

	//! Query for the collection of all running servers.
	virtual HRESULT COMCALL RunningServers(SAFEARRAY** ppServers);

	//! Query for the topics supported by the server.
	virtual HRESULT COMCALL GetServerTopics(BSTR bstrService, SAFEARRAY** ppTopics);

	//! Open a conversation.
	virtual HRESULT COMCALL OpenConversation(BSTR bstrService, BSTR bstrTopic, IDDEConversation** ppIDDEConv);

	//! Query for the collection of all open conversations.
	virtual HRESULT COMCALL Conversations(IDDEConversations** ppIDDEConvs);

	//! Request an item in CF_TEXT format.
	virtual HRESULT COMCALL RequestTextItem(BSTR bstrService, BSTR bstrTopic, BSTR bstrItem, BSTR* pbstrValue);

	//! Poke an item provided in CF_TEXT format.
	virtual HRESULT COMCALL PokeTextItem(BSTR bstrService, BSTR bstrTopic, BSTR bstrItem, BSTR bstrValue);

	//! Execute a command provided in CF_TEXT format.
	virtual HRESULT COMCALL ExecuteTextCommand(BSTR bstrService, BSTR bstrTopic, BSTR bstrCommand);

	//
	// Other methods.
	//

	//! Aquire the singleton DDE Client.
	static DDE::ClientPtr DDEClient();

private:
	//
	// Members.
	//
	DDE::ClientPtr	m_pDDEClient;		//!< The underlying DDE client.
};

#endif // COMDDECLIENT_HPP
