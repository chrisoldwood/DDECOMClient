////////////////////////////////////////////////////////////////////////////////
//! \file   TestDDEConversation.cpp
//! \brief  The unit tests for the DDEConversation class.
//! \author Chris Oldwood

#include "Common.hpp"
#include <Core/UnitTest.hpp>
#include <WCL/ComPtr.hpp>
#include <WCL/ComStr.hpp>
#include <WCL/VariantBool.hpp>
#import "../DDECOMClient.tlb" raw_interfaces_only no_smart_pointers

TEST_SET(DDEConversation)
{
	typedef WCL::ComPtr<DDECOMClientLib::IDDEClient> IDDEClientPtr;
	typedef WCL::ComPtr<DDECOMClientLib::IDDEConversation> IDDEConversationPtr;

TEST_CASE("closed by default")
{
	IDDEConversationPtr pConv(__uuidof(DDECOMClientLib::DDEConversation));

	VARIANT_BOOL vbOpen = VARIANT_FALSE;

	TEST_TRUE(pConv->IsOpen(&vbOpen) == S_OK);
	TEST_TRUE(IsFalse(vbOpen));
}
TEST_CASE_END

TEST_CASE("service and topic name empty by default")
{
	IDDEConversationPtr pConv(__uuidof(DDECOMClientLib::DDEConversation));

	WCL::ComStr bstrService2, bstrTopic2;

	TEST_TRUE(pConv->get_Service(AttachTo(bstrService2)) == S_OK);
	TEST_TRUE(pConv->get_Topic(AttachTo(bstrTopic2)) == S_OK);

	TEST_TRUE(wcscmp(L"", bstrService2.Get()) == 0);
	TEST_TRUE(wcscmp(L"", bstrTopic2.Get()) == 0);
}
TEST_CASE_END

TEST_CASE("conversation can be opened via the client")
{
	IDDEClientPtr pDDEClient(__uuidof(DDECOMClientLib::DDEClient));
	IDDEConversationPtr pConv;

	WCL::ComStr bstrService(TXT("PROGMAN"));
	WCL::ComStr bstrTopic(TXT("PROGMAN"));

	TEST_TRUE(pDDEClient->OpenConversation(bstrService.Get(), bstrTopic.Get(), AttachTo(pConv)) == S_OK);

	VARIANT_BOOL vbOpen = VARIANT_FALSE;

	TEST_TRUE(pConv->IsOpen(&vbOpen) == S_OK);
	TEST_TRUE(IsTrue(vbOpen));
}
TEST_CASE_END

TEST_CASE("service and topic name set after being opened")
{
	IDDEClientPtr pDDEClient(__uuidof(DDECOMClientLib::DDEClient));
	IDDEConversationPtr pConv;

	WCL::ComStr bstrService(TXT("PROGMAN"));
	WCL::ComStr bstrTopic(TXT("PROGMAN"));

	pDDEClient->OpenConversation(bstrService.Get(), bstrTopic.Get(), AttachTo(pConv));

	WCL::ComStr bstrService2, bstrTopic2;

	TEST_TRUE(pConv->get_Service(AttachTo(bstrService2)) == S_OK);
	TEST_TRUE(pConv->get_Topic(AttachTo(bstrTopic2)) == S_OK);

	TEST_TRUE(wcscmp(bstrService.Get(), bstrService2.Get()) == 0);
	TEST_TRUE(wcscmp(bstrTopic.Get(), bstrTopic2.Get()) == 0);
}
TEST_CASE_END

TEST_CASE("open fails when service and topic have not been set")
{
	IDDEConversationPtr pConv(__uuidof(DDECOMClientLib::DDEConversation));

	TEST_TRUE(FAILED(pConv->Open()));
}
TEST_CASE_END

TEST_CASE("conversation can be directly opened after setting service and topic name")
{
	IDDEConversationPtr pConv(__uuidof(DDECOMClientLib::DDEConversation));

	WCL::ComStr bstrService(TXT("Shell"));
	WCL::ComStr bstrTopic(TXT("AppProperties"));

	TEST_TRUE(pConv->put_Service(bstrService.Get()) == S_OK);
	TEST_TRUE(pConv->put_Topic(bstrTopic.Get()) == S_OK);
	TEST_TRUE(pConv->Open() == S_OK);
}
TEST_CASE_END

TEST_CASE("CF_TEXT item can be requested")
{
	IDDEClientPtr pDDEClient(__uuidof(DDECOMClientLib::DDEClient));
	IDDEConversationPtr pConv;

	WCL::ComStr bstrService(TXT("PROGMAN"));
	WCL::ComStr bstrTopic(TXT("PROGMAN"));

	pDDEClient->OpenConversation(bstrService.Get(), bstrTopic.Get(), AttachTo(pConv));

	WCL::ComStr bstrItem(TXT("Accessories"));
	WCL::ComStr bstrValue;

	TEST_TRUE(pConv->RequestTextItem(bstrItem.Get(), AttachTo(bstrValue)) == S_OK);
	TEST_TRUE(wcsstr(bstrValue.Get(), L"Notepad") != nullptr);
}
TEST_CASE_END

TEST_CASE("conversation is not open after being closed")
{
	IDDEClientPtr pDDEClient(__uuidof(DDECOMClientLib::DDEClient));
	IDDEConversationPtr pConv;

	WCL::ComStr bstrService(TXT("PROGMAN"));
	WCL::ComStr bstrTopic(TXT("PROGMAN"));

	pDDEClient->OpenConversation(bstrService.Get(), bstrTopic.Get(), AttachTo(pConv));

	VARIANT_BOOL vbOpen = VARIANT_TRUE;

	TEST_TRUE(pConv->Close() == S_OK);
	TEST_TRUE(pConv->IsOpen(&vbOpen) == S_OK);
	TEST_TRUE(IsFalse(vbOpen));
}
TEST_CASE_END

TEST_CASE("conversation can be opened on a different service and topic after being closed first")
{
	IDDEConversationPtr pConv(__uuidof(DDECOMClientLib::DDEConversation));

	WCL::ComStr bstrService(TXT("Shell"));
	WCL::ComStr bstrTopic(TXT("AppProperties"));

	pConv->put_Service(bstrService.Get());
	pConv->put_Topic(bstrTopic.Get());
	pConv->Open();
	pConv->Close();

	WCL::ComStr bstrService2(TXT("PROGMAN"));
	WCL::ComStr bstrTopic2(TXT("PROGMAN"));

	TEST_TRUE(pConv->put_Service(bstrService2.Get()) == S_OK);
	TEST_TRUE(pConv->put_Topic(bstrTopic2.Get()) == S_OK);
	TEST_TRUE(pConv->Open() == S_OK);
}
TEST_CASE_END

TEST_CASE("service and topic cannot be changed when the conversation is already open")
{
	IDDEConversationPtr pConv(__uuidof(DDECOMClientLib::DDEConversation));

	WCL::ComStr bstrService(TXT("Shell"));
	WCL::ComStr bstrTopic(TXT("AppProperties"));

	pConv->put_Service(bstrService.Get());
	pConv->put_Topic(bstrTopic.Get());
	pConv->Open();

	WCL::ComStr bstrService2(TXT("PROGMAN"));
	WCL::ComStr bstrTopic2(TXT("PROGMAN"));

	TEST_TRUE(FAILED(pConv->put_Service(bstrService2.Get())));
	TEST_TRUE(FAILED(pConv->put_Topic(bstrTopic2.Get())));
	TEST_TRUE(FAILED(pConv->Open()));
}
TEST_CASE_END

}
TEST_SET_END
