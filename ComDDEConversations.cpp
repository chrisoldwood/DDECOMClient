////////////////////////////////////////////////////////////////////////////////
//! \file   ComDDEConversations.cpp
//! \brief  The ComDDEConversations class definition.
//! \author Chris Oldwood

#include "Common.hpp"
#include "ComDDEConversations.hpp"
#include "ComDDEConversation.hpp"
#include "ComDDEConvEnum.hpp"

////////////////////////////////////////////////////////////////////////////////
//! Default constructor.

ComDDEConversations::ComDDEConversations(const DDECltConvs& vConvs)
	: COM::IDispatchImpl<ComDDEConversations>(IID_IDDEConversations)
	, m_vDDEConvs(vConvs.begin(), vConvs.end())
{
	
}

////////////////////////////////////////////////////////////////////////////////
//! Destructor.

ComDDEConversations::~ComDDEConversations()
{
}

////////////////////////////////////////////////////////////////////////////////
//! Get the collection count.

HRESULT COMCALL ComDDEConversations::get_Count(long* plCount)
{
	HRESULT hr = E_FAIL;

	try
	{
		// Validate parameters.
		if (plCount == nullptr)
			throw WCL::ComException(E_POINTER, TXT("plCount is NULL"));

		// Return the count.
		*plCount = m_vDDEConvs.size();

		hr = S_OK;
	}
	COM_CATCH(hr)

	return hr;
}

////////////////////////////////////////////////////////////////////////////////
//! Get the enumerator for the collection.

HRESULT COMCALL ComDDEConversations::get__NewEnum(IUnknown** ppEnum)
{
	HRESULT hr = E_FAIL;

	try
	{
		// Type shorthands.
		typedef Core::IFacePtr<IEnumVARIANT> IEnumVARIANTPtr;

		// Validate parameters.
		if (ppEnum == nullptr)
			throw WCL::ComException(E_POINTER, TXT("ppEnum is NULL"));

		// Reset output parameters.
		*ppEnum = nullptr;

		IEnumVARIANTPtr pEnum = IEnumVARIANTPtr(new ComDDEConvEnum(m_vDDEConvs), true);

		// Return value.
		*ppEnum = pEnum.Detach();

		hr = S_OK;
	}
	COM_CATCH(hr)

	return hr;
}

////////////////////////////////////////////////////////////////////////////////
//! Get an item from the collection by index. The index is assumed to be
//! 0-based as that is the convention for VBScript and WMI.

HRESULT COMCALL ComDDEConversations::get_Item(long nIndex, IDDEConversation** ppIDDEConv)
{
	HRESULT hr = E_FAIL;

	try
	{
		// Type shorthands.
		typedef Core::IFacePtr<IDDEConversation> IDDEConversationPtr;

		// Validate parameters.
		if (ppIDDEConv == nullptr)
			throw WCL::ComException(E_POINTER, TXT("ppIDDEConv is NULL"));

		// Reset output parameters.
		*ppIDDEConv = nullptr;

		// Avoid "signed/unsigned mismatch" errors.
		long nSize = m_vDDEConvs.size();
	
		// Validate input parameters.
		if ( (nIndex < 0) || (nIndex >= nSize) )
			throw WCL::ComException(E_INVALIDARG, TXT("nIndex is out of range"));

		IDDEConversationPtr pComConv = IDDEConversationPtr(new ComDDEConversation(m_vDDEConvs[nIndex]), true);

		// Return value.
		*ppIDDEConv = pComConv.Detach();

		hr = S_OK;
	}
	COM_CATCH(hr)

	return hr;
}
