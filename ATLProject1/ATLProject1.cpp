// ATLProject1.cpp: Implementierung von WinMain


#include "pch.h"
#include "framework.h"
#include "resource.h"
#include "ATLProject1_i.h"
#include "xdlldata.h"


using namespace ATL;


class CATLProject1Module : public ATL::CAtlExeModuleT< CATLProject1Module >
{
public :
	DECLARE_LIBID(LIBID_ATLProject1Lib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_ATLPROJECT1, "{97bb14db-8756-4cf1-8f69-d7057e6a27b0}")
};

CATLProject1Module _AtlModule;



//
extern "C" int WINAPI _tWinMain(HINSTANCE /*hInstance*/, HINSTANCE /*hPrevInstance*/,
								LPTSTR /*lpCmdLine*/, int nShowCmd)
{
	return _AtlModule.WinMain(nShowCmd);
}

