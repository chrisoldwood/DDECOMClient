////////////////////////////////////////////////////////////////////////////////
//! \file   ComDDEConvEnum.cpp
//! \brief  The ComDDEConvEnum class definition.
//! \author Chris Oldwood

#include "Common.hpp"
#include "ComDDEConvEnum.hpp"
#include "ComDDEConversation.hpp"

////////////////////////////////////////////////////////////////////////////////
//! Default constructor.

ComDDEConvEnum::ComDDEConvEnum(const DDECltConvs& vConvs)
	: m_vDDEConvs(vConvs)
	, m_itConv(m_vDDEConvs.begin())
{
}

////////////////////////////////////////////////////////////////////////////////
//! Copy constructor.

ComDDEConvEnum::ComDDEConvEnum(const ComDDEConvEnum& oEnum)
	: m_vDDEConvs(oEnum.m_vDDEConvs)
	, m_itConv(m_vDDEConvs.begin())
{
	// Replicate the source iterator position.
	size_t nIndex = std::distance(oEnum.m_vDDEConvs.begin(), oEnum.m_itConv);

	std::advance(m_itConv, nIndex);
}

////////////////////////////////////////////////////////////////////////////////
//! Destructor.

ComDDEConvEnum::~ComDDEConvEnum()
{
}

////////////////////////////////////////////////////////////////////////////////
//! Retrieve a number of items from the collection.

HRESULT COMCALL ComDDEConvEnum::Next(ULONG nCount, VARIANT* avItems, ULONG* plFetched)
{
	HRESULT hr = E_FAIL;

	try
	{
		// Type shorthands.
		typedef WCL::IFacePtr<IDDEConversation> IDDEConversationPtr;

		// Validate parameters.
		if (avItems == nullptr)
			throw WCL::ComException(E_POINTER, TXT("avItems is NULL"));

		// Reset output parameters.
		for (size_t i = 0; i < nCount; ++i)
			::VariantInit(&avItems[i]);

		if (plFetched != nullptr)
			*plFetched = 0;

		// Calculate how many we can return.
		size_t nAvail   = std::distance(m_itConv, static_cast<ConstIter>(m_vDDEConvs.end()));
		size_t nFetched = std::min<size_t>(nAvail, nCount);

		// Fill in the return buffer.
		VARIANT*  itVarStart = avItems;
		VARIANT*  itVarEnd   = itVarStart+nFetched;
		
		for (VARIANT* itVar = itVarStart; itVar < itVarEnd; ++itVar, ++m_itConv)
		{
			ASSERT(m_itConv != m_vDDEConvs.end());

			IDDEConversationPtr pComConv = IDDEConversationPtr(new ComDDEConversation(*m_itConv), true);

			V_VT(itVar)       = VT_DISPATCH;
			V_DISPATCH(itVar) = pComConv.Detach();
		}

		// Set the number fetched, if requested.
		if (plFetched != nullptr)
			*plFetched = nFetched;

		hr = (nFetched == nCount) ? S_OK : S_FALSE;
	}
	COM_CATCH(hr)

	return hr;
}

////////////////////////////////////////////////////////////////////////////////
//! Skip a number of items in the sequence.

HRESULT COMCALL ComDDEConvEnum::Skip(ULONG nCount)
{
	HRESULT hr = E_FAIL;

	try
	{
		// Advance the iterator.
		while ( (m_itConv != m_vDDEConvs.end()) && (nCount > 0) )
		{
			++m_itConv;
			--nCount;
		}
			
		hr = (nCount == 0) ? S_OK : S_FALSE;
	}
	COM_CATCH(hr)

	return hr;
}

////////////////////////////////////////////////////////////////////////////////
//! Reset the iterator to the start of the sequence.

HRESULT COMCALL ComDDEConvEnum::Reset()
{
	HRESULT hr = E_FAIL;

	try
	{
		// Reset the iterator.
		m_itConv = m_vDDEConvs.begin();

		hr = S_OK;
	}
	COM_CATCH(hr)

	return hr;
}

////////////////////////////////////////////////////////////////////////////////
//! Create a copy of the enumerator.

HRESULT COMCALL ComDDEConvEnum::Clone(IEnumVARIANT** ppEnum)
{
	HRESULT hr = E_FAIL;

	try
	{
		// Type shorthands.
		typedef WCL::IFacePtr<IEnumVARIANT> IEnumVARIANTPtr;

		// Validate parameters.
		if (ppEnum == nullptr)
			throw WCL::ComException(E_POINTER, TXT("ppEnum is NULL"));

		// Reset output parameters.
		*ppEnum = nullptr;

		IEnumVARIANTPtr pEnum = IEnumVARIANTPtr(new ComDDEConvEnum(*this), true);

		// Return value.
		*ppEnum = pEnum.Detach();

		hr = S_OK;
	}
	COM_CATCH(hr)

	return hr;
}
