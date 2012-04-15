////////////////////////////////////////////////////////////////////////////////
//! \file   TestDDEConversations.cpp
//! \brief  The unit tests for the DDEConversations class.
//! \author Chris Oldwood

#include "stdafx.h"
#include <Core/UnitTest.hpp>
#include <WCL/ComPtr.hpp>
#include <WCL/ComStr.hpp>
#include <WCL/VariantBool.hpp>
#import "../DDECOMClient.tlb" raw_interfaces_only no_smart_pointers

TEST_SET(DDEConversations)
{
	typedef WCL::ComPtr<DDECOMClientLib::IDDEClient> IDDEClientPtr;
	typedef WCL::ComPtr<DDECOMClientLib::IDDEConversation> IDDEConversationPtr;
	typedef WCL::ComPtr<DDECOMClientLib::IDDEConversations> IDDEConversationsPtr;
	typedef WCL::IFacePtr<IUnknown> IUnknownPtr;
	typedef WCL::ComPtr<IEnumVARIANT> IEnumVARIANTPtr;

	IDDEClientPtr pDDEClient(__uuidof(DDECOMClientLib::DDEClient));

	WCL::ComStr bstrService1(TXT("PROGMAN"));
	WCL::ComStr bstrTopic1(TXT("PROGMAN"));
	IDDEConversationPtr pConv1;

	HRESULT result1 = pDDEClient->OpenConversation(bstrService1.Get(), bstrTopic1.Get(), AttachTo(pConv1));

	WCL::ComStr bstrService2(TXT("Shell"));
	WCL::ComStr bstrTopic2(TXT("AppProperties"));
	IDDEConversationPtr pConv2;

	HRESULT result2 = pDDEClient->OpenConversation(bstrService2.Get(), bstrTopic2.Get(), AttachTo(pConv2));

	ASSERT(result1 == S_OK);
	ASSERT(result2 == S_OK);

TEST_CASE("client can be queried for collection of open conversations")
{
	IDDEConversationsPtr pConvs;

	TEST_TRUE(pDDEClient->Conversations(AttachTo(pConvs)) == S_OK);
}
TEST_CASE_END

TEST_CASE("count of conversations can be queried from collection")
{
	IDDEConversationsPtr pConvs;

	pDDEClient->Conversations(AttachTo(pConvs));

	long lCount = 0;

	TEST_TRUE(pConvs->get_Count(&lCount) == S_OK);
	TEST_TRUE(lCount == 2);
}
TEST_CASE_END

TEST_CASE("each conversation can be queried from the collection by index")
{
	IDDEConversationsPtr pConvs;

	pDDEClient->Conversations(AttachTo(pConvs));

	IDDEConversationPtr pConv3, pConv4;

	TEST_TRUE(pConvs->get_Item(0, AttachTo(pConv3)) == S_OK);
	TEST_TRUE(pConvs->get_Item(1, AttachTo(pConv4)) == S_OK);

	WCL::ComStr bstrService3, bstrService4;

	pConv3->get_Service(AttachTo(bstrService3));
	pConv4->get_Service(AttachTo(bstrService4));

	TEST_TRUE(wcscmp(bstrService3.Get(), bstrService4.Get()) != 0);
	TEST_TRUE((wcscmp(bstrService3.Get(), bstrService1.Get()) == 0) || (wcscmp(bstrService3.Get(), bstrService2.Get()) == 0));
	TEST_TRUE((wcscmp(bstrService4.Get(), bstrService1.Get()) == 0) || (wcscmp(bstrService4.Get(), bstrService2.Get()) == 0));
}
TEST_CASE_END

	IDDEConversationsPtr pConvs;
	long lCount = 0;

	pDDEClient->Conversations(AttachTo(pConvs));
	pConvs->get_Count(&lCount);

	ASSERT(lCount == 2);

TEST_CASE("collection supports standard enumerator")
{
	IUnknownPtr pUnknown;

	TEST_TRUE(pConvs->get__NewEnum(AttachTo(pUnknown)) == S_OK);

	IEnumVARIANTPtr pEnum(pUnknown);
}
TEST_CASE_END

TEST_CASE("cannot skip over more elements than are in the collection")
{
	IUnknownPtr pUnknown;

	pConvs->get__NewEnum(AttachTo(pUnknown));

	IEnumVARIANTPtr pEnum(pUnknown);

	TEST_TRUE(pEnum->Skip(100) == S_FALSE);
}
TEST_CASE_END

TEST_CASE("enumerator can skip items")
{
	IUnknownPtr pUnknown;

	pConvs->get__NewEnum(AttachTo(pUnknown));

	IEnumVARIANTPtr pEnum(pUnknown);

	TEST_TRUE(pEnum->Skip(0) == S_OK);
	TEST_TRUE(pEnum->Skip(1) == S_OK);
	TEST_TRUE(pEnum->Skip(1) == S_OK);
	TEST_TRUE(pEnum->Skip(1) == S_FALSE);
}
TEST_CASE_END

TEST_CASE("enumerator can be cloned with position maintained")
{
	IUnknownPtr pUnknown;

	pConvs->get__NewEnum(AttachTo(pUnknown));

	IEnumVARIANTPtr pEnum1(pUnknown);

	pEnum1->Skip(1);

	IEnumVARIANTPtr pEnum2;

	TEST_TRUE(pEnum1->Clone(AttachTo(pEnum2)) == S_OK);

	TEST_TRUE(pEnum1->Skip(1) == S_OK);
	TEST_TRUE(pEnum2->Skip(1) == S_OK);
	TEST_TRUE(pEnum1->Skip(1) == S_FALSE);
	TEST_TRUE(pEnum2->Skip(1) == S_FALSE);
}
TEST_CASE_END

TEST_CASE("enumerator can be reset")
{
	IUnknownPtr pUnknown;

	pConvs->get__NewEnum(AttachTo(pUnknown));

	IEnumVARIANTPtr pEnum(pUnknown);

	pEnum->Skip(2);

	ASSERT(pEnum->Skip(1) == S_FALSE);

	TEST_TRUE(pEnum->Reset() == S_OK);
	TEST_TRUE(pEnum->Skip(2) == S_OK);
	TEST_TRUE(pEnum->Skip(1) == S_FALSE);
}
TEST_CASE_END
/*
TEST_CASE("enumerator can|cannot (?) skip 0 items when sequence end reached")
{
	IUnknownPtr pUnknown;

	pConvs->get__NewEnum(AttachTo(pUnknown));

	IEnumVARIANTPtr pEnum(pUnknown);

	pEnum->Skip(2);

	TEST_TRUE(pEnum->Skip(1) == S_FALSE);
	TEST_TRUE(pEnum->Skip(0) == S_FALSE);
}
TEST_CASE_END
*/
}
TEST_SET_END
