// Test.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include <iomanip>
#include <WCL/AutoCom.hpp>
#include <WCL/ComPtr.hpp>
#include <WCL/SafeVector.hpp>
#include <WCL/ComStr.hpp>

#import "../DDECOMClient.tlb" raw_interfaces_only no_smart_pointers

void TestDDEClient()
{
	try
	{
		typedef WCL::ComPtr<DDECOMClientLib::IDDEClient> IDDEClientPtr;
		typedef Core::IFacePtr<DDECOMClientLib::IDDEConversation> IDDEConversationPtr;
		typedef WCL::SafeVector<VARIANT> VariantArray;
		typedef VariantArray::const_iterator  CIter;

		IDDEClientPtr pDDEClient(__uuidof(DDECOMClientLib::DDEClient));

		SAFEARRAY* pServersArray;

		HRESULT hr = pDDEClient->RunningServers(&pServersArray);

		if (FAILED(hr))
			throw WCL::ComException(hr, pDDEClient, "");

		VariantArray avtServers(pServersArray);

		std::cout << "Running Servers & Topics: " << avtServers.Size() << std::endl;

		for (CIter itServer = avtServers.begin(); itServer != avtServers.end(); ++itServer)
		{
			std::wcout << V_BSTR(itServer) << std::endl;

			SAFEARRAY* pTopicsArray;

			HRESULT hr = pDDEClient->GetServerTopics(V_BSTR(itServer), &pTopicsArray);

			if (FAILED(hr))
				throw WCL::ComException(hr, pDDEClient, "");

			VariantArray avtTopics(pTopicsArray);

			for (CIter itTopic = avtTopics.begin(); itTopic != avtTopics.end(); ++itTopic)
				std::wcout << L"\t" << V_BSTR(itTopic) << std::endl;
		}

		std::cout << std::endl;
		std::cout << TXT("Requesting item - PROGMAN|PROGMAN!Accessories") << std::endl;

		WCL::ComStr bstrService(L"PROGMAN");
		WCL::ComStr bstrTopic(L"PROGMAN");
		WCL::ComStr bstrItem("Accessories");
		WCL::ComStr bstrValue;

		hr = pDDEClient->RequestTextItem(bstrService.Get(), bstrTopic.Get(), bstrItem.Get(), AttachTo(bstrValue));

		if (FAILED(hr))
			throw WCL::ComException(hr, pDDEClient, "");

		std::wcout << bstrValue.Get() << std::endl;
		std::cout  << std::endl;
		std::cout  << TXT("Opening conversation - PROGMAN|PROGMAN") << std::endl;

		IDDEConversationPtr pConv;

		hr = pDDEClient->OpenConversation(bstrService.Get(), bstrTopic.Get(), AttachTo(pConv));

		if (FAILED(hr))
			throw WCL::ComException(hr, pDDEClient, "");

		std::cout << TXT("Requesting item - Accessories") << std::endl;

		hr = pConv->RequestTextItem(bstrItem.Get(), AttachTo(bstrValue));

		if (FAILED(hr))
			throw WCL::ComException(hr, pDDEClient, "");

		std::wcout << bstrValue.Get() << std::endl;
		std::cout  << std::endl;
	}
	catch (const std::exception& e)
	{
		std::cout << e.what() << std::endl;
	}
}

int _tmain(int /*argc*/, _TCHAR* /*argv*/[])
{
	WCL::AutoCom oCom(COINIT_APARTMENTTHREADED);
//	WCL::AutoCom oCom(COINIT_MULTITHREADED);

	TestDDEClient();

	return EXIT_SUCCESS;
}
