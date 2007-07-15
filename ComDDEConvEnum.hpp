////////////////////////////////////////////////////////////////////////////////
//! \file   ComDDEConvEnum.hpp
//! \brief  The ComDDEConvEnum class declaration.
//! \author Chris Oldwood

// Check for previous inclusion
#ifndef COMDDECONVENUM_HPP
#define COMDDECONVENUM_HPP

#include <COM/IDispatchImpl.hpp>
#include <NCL/DDECltConvPtr.hpp>
#include <vector>

#if _MSC_VER > 1000
#pragma once
#endif

// The conversation container type.
typedef std::vector<DDE::CltConvPtr> DDECltConvs;

////////////////////////////////////////////////////////////////////////////////
//! The enumerator for the DDE conversations.

class ComDDEConvEnum : public COM::ObjectBase<IEnumVARIANT>
{
public:
	//! Default constructor.
	ComDDEConvEnum(const DDECltConvs& vConvs);

	//! Copy constructor.
	ComDDEConvEnum(const ComDDEConvEnum& oEnum);

	//! Destructor.
	virtual ~ComDDEConvEnum();
	
	DEFINE_INTERFACE_TABLE(ComDDEConvEnum)
		IMPLEMENT_INTERFACE(IID_IEnumVARIANT, IEnumVARIANT)
	END_INTERFACE_TABLE()

	IMPLEMENT_IUNKNOWN()

	//
	// IEnumVARIANT methods.
	//

	//! Retrieve a number of items from the collection.
	virtual HRESULT COMCALL Next(ULONG nCount, VARIANT* avItems, ULONG* plFetched);

	//! Skip a number of items in the sequence.
	virtual HRESULT COMCALL Skip(ULONG nCount);

	//! Reset the iterator to the start of the sequence.
	virtual HRESULT COMCALL Reset();

	//! Create a copy of the enumerator.
	virtual HRESULT COMCALL Clone(IEnumVARIANT** ppEnum);

private:
	// Type shorthands.
	typedef DDECltConvs::const_iterator ConstIter;

	//
	// Members.
	//
	DDECltConvs	m_vDDEConvs;	//!< The collection of underlying conversations.
	ConstIter	m_itConv;		//!< The underlying collection iterator.
};

#endif // COMDDECONVENUM_HPP
