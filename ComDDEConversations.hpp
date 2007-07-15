////////////////////////////////////////////////////////////////////////////////
//! \file   ComDDEConversations.hpp
//! \brief  The ComDDEConversations class declaration.
//! \author Chris Oldwood

// Check for previous inclusion
#ifndef COMDDECONVERSATIONS_HPP
#define COMDDECONVERSATIONS_HPP

#include <COM/IDispatchImpl.hpp>
#include <NCL/DDECltConvPtr.hpp>
#include <vector>

#if _MSC_VER > 1000
#pragma once
#endif

// The conversation container type.
typedef std::vector<DDE::CltConvPtr> DDECltConvs;

////////////////////////////////////////////////////////////////////////////////
//! A collection of DDE conversations.

class ComDDEConversations : public COM::ObjectBase<IDDEConversations>, COM::IDispatchImpl<ComDDEConversations>
{
public:
	//! Default constructor.
	ComDDEConversations(const DDECltConvs& vConvs);

	//! Destructor.
	virtual ~ComDDEConversations();
	
	DEFINE_INTERFACE_TABLE(IDDEConversations)
		IMPLEMENT_INTERFACE(IID_IDDEConversations, IDDEConversations)
		IMPLEMENT_INTERFACE(IID_IDispatch,         IDDEConversations)
	END_INTERFACE_TABLE()

	IMPLEMENT_IUNKNOWN()
	IMPLEMENT_IDISPATCH(ComDDEConversations)

	//
	// IDDEConversations methods.
	//

	//! Get the collection count.
	virtual HRESULT COMCALL get_Count(long* plCount);

	//! Get the enumerator for the collection.
	virtual HRESULT COMCALL get__NewEnum(IUnknown** ppEnum);

	//! Get an item from the collection by index.
	virtual HRESULT COMCALL get_Item(long nIndex, IDDEConversation** ppIDDEConv);

private:
	//
	// Members.
	//
	DDECltConvs	m_vDDEConvs;	//!< The collection of underlying conversations.
};

#endif // COMDDECONVERSATIONS_HPP
