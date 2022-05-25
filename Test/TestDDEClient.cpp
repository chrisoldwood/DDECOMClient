////////////////////////////////////////////////////////////////////////////////
//! \file   TestDDEClient.cpp
//! \brief  The unit tests for the DDEClient class.
//! \author Chris Oldwood

#include "Common.hpp"
#include <Core/UnitTest.hpp>
#include <WCL/ComPtr.hpp>
#include <WCL/VariantVector.hpp>
#include <WCL/ComStr.hpp>
#import "../DDECOMClient.tlb" raw_interfaces_only no_smart_pointers
#include <algorithm>

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

TEST_SET(DDEClient)
{
	typedef WCL::ComPtr<DDECOMClientLib::IDDEClient> IDDEClientPtr;
	typedef WCL::VariantVector<VARIANT> VariantArray;

TEST_CASE("client object can be instantiated")
{
	IDDEClientPtr pDDEClient(__uuidof(DDECOMClientLib::DDEClient));

	TEST_TRUE(pDDEClient.get() != nullptr);
}
TEST_CASE_END

TEST_CASE("running servers can be retrieved")
{
	IDDEClientPtr pDDEClient(__uuidof(DDECOMClientLib::DDEClient));

	SAFEARRAY* pServersArray = nullptr;

	TEST_TRUE(pDDEClient->RunningServers(&pServersArray) == S_OK);

	VariantArray avtServers(pServersArray);

	TEST_TRUE(avtServers.Size() >= 1);
	TEST_TRUE(std::find_if(avtServers.begin(), avtServers.end(), Compare(L"PROGMAN")) != avtServers.end());
}
TEST_CASE_END

TEST_CASE("supported topics for a service can be retrieved")
{
	IDDEClientPtr pDDEClient(__uuidof(DDECOMClientLib::DDEClient));

	SAFEARRAY* pTopicsArray = nullptr;

	TEST_TRUE(pDDEClient->GetServerTopics(WCL::ComStr(TXT("PROGMAN")).Get(), &pTopicsArray) == S_OK);

	VariantArray avtTopics(pTopicsArray);

	TEST_TRUE(std::find_if(avtTopics.begin(), avtTopics.end(), Compare(L"PROGMAN")) != avtTopics.end());
}
TEST_CASE_END

TEST_CASE("single text item for a service and topic can be retrieved")
{
	IDDEClientPtr pDDEClient(__uuidof(DDECOMClientLib::DDEClient));

	WCL::ComStr bstrService(TXT("PROGMAN"));
	WCL::ComStr bstrTopic(TXT("PROGMAN"));
	WCL::ComStr bstrItem(TXT("Accessories"));
	WCL::ComStr bstrValue;

	HRESULT result = pDDEClient->RequestTextItem(bstrService.Get(), bstrTopic.Get(), bstrItem.Get(), AttachTo(bstrValue));

	TEST_TRUE(result == S_OK);
	TEST_TRUE(wcsstr(bstrValue.Get(), L"Accessories") != nullptr);
}
TEST_CASE_END

}
TEST_SET_END
