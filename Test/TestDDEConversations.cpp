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

////////////////////////////////////////////////////////////////////////////////
//! The unit tests for the DDEConversations class.

void TestDDEConversations()
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

	TEST_TRUE(pDDEClient->OpenConversation(bstrService1.Get(), bstrTopic1.Get(), AttachTo(pConv1)) == S_OK);

	WCL::ComStr bstrService2(TXT("Shell"));
	WCL::ComStr bstrTopic2(TXT("AppProperties"));
	IDDEConversationPtr pConv2;

	TEST_TRUE(pDDEClient->OpenConversation(bstrService2.Get(), bstrTopic2.Get(), AttachTo(pConv2)) == S_OK);

	IDDEConversationsPtr pConvs;

	TEST_TRUE(pDDEClient->Conversations(AttachTo(pConvs)) == S_OK);

	long lCount = 0;

	TEST_TRUE(pConvs->get_Count(&lCount) == S_OK);
	TEST_TRUE(lCount == 2);

	IDDEConversationPtr pConv3, pConv4;

	TEST_TRUE(pConvs->get_Item(0, AttachTo(pConv3)) == S_OK);
	TEST_TRUE(pConvs->get_Item(1, AttachTo(pConv4)) == S_OK);

	WCL::ComStr bstrService3, bstrService4;

	TEST_TRUE(pConv3->get_Service(AttachTo(bstrService3)) == S_OK);
	TEST_TRUE(pConv4->get_Service(AttachTo(bstrService4)) == S_OK);

	TEST_TRUE(wcscmp(bstrService3.Get(), bstrService4.Get()) != 0);
	TEST_TRUE((wcscmp(bstrService3.Get(), bstrService1.Get()) == 0) || (wcscmp(bstrService3.Get(), bstrService2.Get()) == 0));
	TEST_TRUE((wcscmp(bstrService4.Get(), bstrService1.Get()) == 0) || (wcscmp(bstrService4.Get(), bstrService2.Get()) == 0));

	IUnknownPtr pUnknown;

	TEST_TRUE(pConvs->get__NewEnum(AttachTo(pUnknown)) == S_OK);

	IEnumVARIANTPtr pEnum1(pUnknown);

	TEST_TRUE(pEnum1->Skip(0) == S_OK);
	TEST_TRUE(pEnum1->Skip(1) == S_OK);

	IEnumVARIANTPtr pEnum2;

	TEST_TRUE(pEnum1->Clone(AttachTo(pEnum2)) == S_OK);
	TEST_TRUE(pEnum1->Skip(1) == S_OK);
	TEST_TRUE(pEnum1->Skip(1) == S_FALSE);

	TEST_TRUE(pEnum2->Skip(1) == S_OK);
	TEST_TRUE(pEnum2->Skip(1) == S_FALSE);
	TEST_TRUE(pEnum2->Reset() == S_OK);
	TEST_TRUE(pEnum2->Skip(2) == S_OK);
}
