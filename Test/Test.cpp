////////////////////////////////////////////////////////////////////////////////
//! \file   Test.cpp
//! \brief  The test harness entry point.
//! \author Chris Oldwood

#include "stdafx.h"
#include <tchar.h>
#include <Core/UnitTest.hpp>
#include <WCL/AutoCom.hpp>

////////////////////////////////////////////////////////////////////////////////
// The test group functions.

extern void TestDDEClient();
extern void TestDDEConversation();
extern void TestDDEConversations();

////////////////////////////////////////////////////////////////////////////////
//! The entry point for the test harness.

int _tmain(int /*argc*/, _TCHAR* /*argv*/[])
{
	TEST_SUITE_BEGIN
	{
		WCL::AutoCom oCom(COINIT_APARTMENTTHREADED);
//		WCL::AutoCom oCom(COINIT_MULTITHREADED);

		TestDDEClient();
		TestDDEConversation();
		TestDDEConversations();

		Core::SetTestRunFinalStatus(true);
	}
	TEST_SUITE_END
}
