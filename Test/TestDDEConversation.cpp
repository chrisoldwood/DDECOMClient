////////////////////////////////////////////////////////////////////////////////
//! \file   TestDDEConversation.cpp
//! \brief  The unit tests for the DDEConversation class.
//! \author Chris Oldwood

#include "stdafx.h"
#include <Core/UnitTest.hpp>
#include <WCL/ComPtr.hpp>
#include <WCL/ComStr.hpp>
#include <WCL/VariantBool.hpp>
#import "../DDECOMClient.tlb" raw_interfaces_only no_smart_pointers

////////////////////////////////////////////////////////////////////////////////
//! The unit tests for the DDEConversation class.

void TestDDEConversation()
{
	typedef WCL::ComPtr<DDECOMClientLib::IDDEClient> IDDEClientPtr;
	typedef WCL::ComPtr<DDECOMClientLib::IDDEConversation> IDDEConversationPtr;

	IDDEConversationPtr pConv(__uuidof(DDECOMClientLib::DDEConversation));

	TEST_TRUE(FAILED(pConv->Open()));
	pConv.Release();

	IDDEClientPtr pDDEClient(__uuidof(DDECOMClientLib::DDEClient));

	WCL::ComStr bstrService1(TXT("PROGMAN"));
	WCL::ComStr bstrTopic1(TXT("PROGMAN"));

	TEST_TRUE(pDDEClient->OpenConversation(bstrService1.Get(), bstrTopic1.Get(), AttachTo(pConv)) == S_OK);

	VARIANT_BOOL vbOpen = VARIANT_FALSE;

	TEST_TRUE(pConv->IsOpen(&vbOpen) == S_OK);
	TEST_TRUE(IsTrue(vbOpen) == true);

	WCL::ComStr bstrService2;

	TEST_TRUE(pConv->get_Service(AttachTo(bstrService2)) == S_OK);
	TEST_TRUE(wcscmp(bstrService1.Get(), bstrService2.Get()) == 0);

	WCL::ComStr bstrTopic2;

	TEST_TRUE(pConv->get_Topic(AttachTo(bstrTopic2)) == S_OK);
	TEST_TRUE(wcscmp(bstrTopic1.Get(), bstrTopic2.Get()) == 0);

	WCL::ComStr bstrItem(TXT("Accessories"));
	WCL::ComStr bstrValue;

	TEST_TRUE(pConv->RequestTextItem(bstrItem.Get(), AttachTo(bstrValue)) == S_OK);
	TEST_TRUE(wcsstr(bstrValue.Get(), L"Windows Explorer") != nullptr);

	TEST_TRUE(pConv->Close() == S_OK);
	TEST_TRUE(pConv->IsOpen(&vbOpen) == S_OK);
	TEST_TRUE(IsTrue(vbOpen) == false);

	WCL::ComStr bstrService3(TXT("Shell"));
	WCL::ComStr bstrTopic3(TXT("AppProperties"));

	bstrService2.Release();
	bstrTopic2.Release();

	TEST_TRUE(pConv->put_Service(bstrService3.Get()) == S_OK);
	TEST_TRUE(pConv->get_Service(AttachTo(bstrService2)) == S_OK);
	TEST_TRUE(wcscmp(bstrService3.Get(), bstrService2.Get()) == 0);

	TEST_TRUE(pConv->put_Topic(bstrTopic3.Get()) == S_OK);
	TEST_TRUE(pConv->get_Topic(AttachTo(bstrTopic2)) == S_OK);
	TEST_TRUE(wcscmp(bstrTopic3.Get(), bstrTopic2.Get()) == 0);

	TEST_TRUE(pConv->Open() == S_OK);
}
