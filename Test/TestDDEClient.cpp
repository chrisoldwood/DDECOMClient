////////////////////////////////////////////////////////////////////////////////
//! \file   TestDDEClient.cpp
//! \brief  The unit tests for the DDEClient class.
//! \author Chris Oldwood

#include "stdafx.h"
#include <Core/UnitTest.hpp>
#include <WCL/ComPtr.hpp>
#include <WCL/SafeVector.hpp>
#include <WCL/ComStr.hpp>
#import "../DDECOMClient.tlb" raw_interfaces_only no_smart_pointers
#include <algorithm>

////////////////////////////////////////////////////////////////////////////////
//! Comparison function for a BSTR inside a VARIANT.

struct Compare
{
	Compare(const wchar_t* psz)
		: m_psz(psz)
	{}

	bool operator()(const VARIANT& vt)
	{
		return wcscmp(m_psz, V_BSTR(&vt)) == 0;
	}

	const wchar_t* m_psz;
};

////////////////////////////////////////////////////////////////////////////////
//! The unit tests for the DDEClient class.

void TestDDEClient()
{
	typedef WCL::ComPtr<DDECOMClientLib::IDDEClient> IDDEClientPtr;

	IDDEClientPtr pDDEClient(__uuidof(DDECOMClientLib::DDEClient));

	TEST_TRUE(pDDEClient.Get() != nullptr);

	typedef WCL::SafeVector<VARIANT> VariantArray;

	SAFEARRAY* pServersArray = nullptr;

	TEST_TRUE(pDDEClient->RunningServers(&pServersArray) == S_OK);

	VariantArray avtServers(pServersArray);

	TEST_TRUE(avtServers.Size() >= 3);
	TEST_TRUE(std::find_if(avtServers.begin(), avtServers.end(), Compare(L"Folders")) != avtServers.end());
	TEST_TRUE(std::find_if(avtServers.begin(), avtServers.end(), Compare(L"Shell")) != avtServers.end());
	TEST_TRUE(std::find_if(avtServers.begin(), avtServers.end(), Compare(L"PROGMAN")) != avtServers.end());

	SAFEARRAY* pTopicsArray = nullptr;

	TEST_TRUE(pDDEClient->GetServerTopics(WCL::ComStr(TXT("Shell")).Get(), &pTopicsArray) == S_OK);

	VariantArray avtTopics(pTopicsArray);

	TEST_TRUE(std::find_if(avtTopics.begin(), avtTopics.end(), Compare(L"AppProperties")) != avtTopics.end());

	WCL::ComStr bstrService(TXT("PROGMAN"));
	WCL::ComStr bstrTopic(TXT("PROGMAN"));
	WCL::ComStr bstrItem(TXT("Accessories"));
	WCL::ComStr bstrValue;

	TEST_TRUE(pDDEClient->RequestTextItem(bstrService.Get(), bstrTopic.Get(), bstrItem.Get(), AttachTo(bstrValue)) == S_OK);
	TEST_TRUE(wcsstr(bstrValue.Get(), L"Internet Explorer") != nullptr);
	TEST_TRUE(wcsstr(bstrValue.Get(), L"Windows Explorer") != nullptr);
}
