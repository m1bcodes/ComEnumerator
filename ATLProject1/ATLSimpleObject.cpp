// ATLSimpleObject.cpp: Implementierung von CATLSimpleObject

#include "pch.h"
#include "ATLSimpleObject.h"
#include <comutil.h>
#include <string>

// CATLSimpleObject


STDMETHODIMP CATLSimpleObject::Reverse2(BSTR Str, BSTR* retval)
{
	_bstr_t bs(Str);
	std::string s = (const char*)bs;
	std::string s1(s.crbegin(), s.crend());
	_bstr_t bs2(s1.c_str());
	*retval = bs2.Detach();
	return S_OK;
}

STDMETHODIMP CATLSimpleObject::Test(FocusPolicy fp, LONG* result)
{
	*result = 42;
	return S_OK;
}
STDMETHODIMP CATLSimpleObject::Test2(FocusPolicy fp, FocusPolicy* fpOut)
{
	return S_OK;
}

STDMETHODIMP CATLSimpleObject::Test3(SAFEARRAY* arrLong, SAFEARRAY* arrEnum, SAFEARRAY** arrDouble)
{
	return S_OK;
}
