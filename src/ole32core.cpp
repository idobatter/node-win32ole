/*
  ole32core.cpp
  This source is independent of node/v8.
*/

#include "ole32core.h"

using namespace std;

BOOL chkerr(BOOL b, char *m, int n, char *f, char *e)
{
  if(b) return b;
  DWORD code = GetLastError();
  WCHAR *buf;
  FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER
    | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
    NULL, code, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
    (LPWSTR)&buf, 0, NULL);
  fprintf(stderr, "ASSERT in %08x module %s(%d) @%s: %s\n", code, m, n, f, e);
  // fwprintf(stderr, L"error: %s", buf); // Some wchars can't see on console.
  char *mbs = wcs2mbs(buf);
  if(!mbs) MessageBoxW(NULL, buf, L"error", MB_ICONEXCLAMATION | MB_OK);
  else fprintf(stderr, "error: %s", mbs);
  free(mbs);
  LocalFree(buf);
  return b;
}

wchar_t *u8s2wcs(char *u8s)
{
  size_t u8len = strlen(u8s);
  size_t wclen = MultiByteToWideChar(CP_UTF8, 0, u8s, u8len, NULL, 0);
  wchar_t *wcs = (wchar_t *)malloc((wclen + 1) * sizeof(wchar_t));
  wclen = MultiByteToWideChar(CP_UTF8, 0, u8s, u8len, wcs, wclen + 1); // + 1
  wcs[wclen] = L'\0';
  return wcs; // ucs2 *** must be free later ***
}

char *wcs2mbs(wchar_t *wcs)
{
  size_t mblen = WideCharToMultiByte(GetACP(), 0,
    (LPCWSTR)wcs, -1, NULL, 0, NULL, NULL);
  char *mbs = (char *)malloc((mblen + 1));
  mblen = WideCharToMultiByte(GetACP(), 0,
    (LPCWSTR)wcs, -1, mbs, mblen, NULL, NULL); // not + 1
  mbs[mblen] = '\0';
  return mbs; // cp932 *** must be free later ***
}

// obsoleted functions

// ShiftJIS -> Unicode -> UTF-8 (without free when use _malloca)
string SJIS2UTF8(string sjis)
{
  if(sjis == "") return "";
  int wlen = MultiByteToWideChar(CP_ACP, 0, sjis.c_str(), -1, NULL, 0);
  WCHAR *wbuf = (WCHAR *)_malloca((wlen + 1) * sizeof(WCHAR));
  if(wbuf == NULL){
    throw "_malloca failed for SJIS to UNICODE";
    return "";
  }
  *wbuf = L'\0';
  if(MultiByteToWideChar(CP_ACP, 0, sjis.c_str(), -1, wbuf, wlen) <= 0){
    throw "can't convert SJIS to UNICODE";
    return "";
  }
  int ulen = WideCharToMultiByte(CP_UTF8, 0, wbuf, -1, NULL, 0, NULL, NULL);
  char *ubuf = (char *)_malloca((ulen + 1) * sizeof(char));
  if(ubuf == NULL){
    throw "_malloca failed for UNICODE to UTF8";
    return "";
  }
  *ubuf = '\0';
  if(WideCharToMultiByte(CP_UTF8, 0, wbuf, -1, ubuf, ulen, NULL, NULL) <= 0){
    throw "can't convert UNICODE to UTF8";
    return "";
  }
  ubuf[ulen] = '\0';
  return ubuf;
}

// UTF8 -> Unicode -> ShiftJIS (without free when use _malloca)
string UTF82SJIS(string utf8)
{
  if(utf8 == "") return "";
  int wlen = MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), -1, NULL, 0);
  WCHAR *wbuf = (WCHAR *)_malloca((wlen + 1) * sizeof(WCHAR));
  if(wbuf == NULL){
    throw "_malloca failed for UTF8 to UNICODE";
    return "";
  }
  *wbuf = L'\0';
  if(MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), -1, wbuf, wlen) <= 0){
    throw "can't convert UTF8 to UNICODE";
    return "";
  }
  int slen = WideCharToMultiByte(CP_ACP, 0, wbuf, -1, NULL, 0, NULL, NULL);
  char *sbuf = (char *)_malloca((slen + 1) * sizeof(char));
  if(sbuf == NULL){
    throw "_malloca failed for UNICODE to SJIS";
    return "";
  }
  *sbuf = '\0';
  if(WideCharToMultiByte(CP_ACP, 0, wbuf, -1, sbuf, slen, NULL, NULL) <= 0){
    throw "can't convert UNICODE to SJIS";
    return "";
  }
  sbuf[slen] = '\0';
  return sbuf;
}

// SjiftJIS -> Unicode (allocate wcs, must free)
WCHAR *SJIS2WCS(string sjis)
{
  if(sjis == ""){
    throw "assigned empty string for SJIS to UNICODE";
    return NULL;
  }
  int wlen = MultiByteToWideChar(CP_ACP, 0, sjis.c_str(), -1, NULL, 0);
  WCHAR *wbuf = new WCHAR[wlen + 1];
  if(wbuf == NULL){
    throw "_malloca failed for SJIS to UNICODE";
    return NULL;
  }
  *wbuf = L'\0';
  if(MultiByteToWideChar(CP_ACP, 0, sjis.c_str(), -1, wbuf, wlen) <= 0){
    delete [] wbuf;
    throw "can't convert SJIS to UNICODE";
    return NULL;
  }
  return wbuf;
}

// Unicode -> ShiftJIS (not free wcs, without free when use _malloca)
string WCS2SJIS(WCHAR *wbuf)
{
  if(wbuf == NULL || wbuf[0] == L'\0') return "";
  int slen = WideCharToMultiByte(CP_ACP, 0, wbuf, -1, NULL, 0, NULL, NULL);
  char *sbuf = (char *)_malloca((slen + 1) * sizeof(char));
  if(sbuf == NULL){
    throw "_malloca failed for UNICODE to SJIS";
    return "";
  }
  *sbuf = '\0';
  if(WideCharToMultiByte(CP_ACP, 0, wbuf, -1, sbuf, slen, NULL, NULL) <= 0){
    throw "can't convert UNICODE to SJIS";
    return "";
  }
  sbuf[slen] = '\0';
  return sbuf;
}

// ShiftJIS -> BSTR (allocate bstr, must free)
BSTR SJIS2BSTR(string str)
{
  BSTR bstr;
  size_t len = str.length();
  WCHAR *wbuf = new WCHAR[len + 1];
  mbstowcs(wbuf, str.c_str(), len);
  wbuf[len] = L'\0';
  bstr = ::SysAllocString(wbuf);
  delete [] wbuf;
  return bstr;
}

// bug ? comment (see old SJIS2UTF8.cpp project)
// BSTR -> ShiftJIS (not free bstr)
string BSTR2SJIS(BSTR bstr)
{
  string str;
  int len = ::SysStringLen(bstr) * 2;
  char *buf = new char[len + 1];
  wcstombs(buf, bstr, len);
  buf[len] = '\0';
  str = buf;
  delete [] buf;
  return str;
}

string OLE32coreException::errorMessage(char *m)
{
  ostringstream oss;
  oss << "OLE error: ";
  if(m) oss << "[" << m << "] ";
  oss << rmsg << endl;
  return oss.str();
}

OCVariant::OCVariant() : next(NULL)
{
  VariantInit(&v);
#ifdef DEBUG
  fprintf(stderr, "--construction-- %08lx %08lx\n", &v, v.vt);
#endif
}

OCVariant::OCVariant(const OCVariant &s) : next(NULL)
{
  VariantInit(&v); // It will be free before copy.
  VariantCopy(&v, (VARIANT *)&s.v);
#ifdef DEBUG
  fprintf(stderr, "--copy construction-- %08lx %08lx\n", &v, v.vt);
#endif
}

OCVariant::OCVariant(long lVal) : next(NULL)
{
  VariantInit(&v);
  v.vt = VT_I4;
  v.lVal = lVal;
#ifdef DEBUG
  fprintf(stderr, "--construction-- %08lx %08lx\n", &v, v.vt);
#endif
}

OCVariant::OCVariant(double dblVal) : next(NULL)
{
  VariantInit(&v);
  v.vt = VT_R8;
  v.dblVal = dblVal;
#ifdef DEBUG
  fprintf(stderr, "--construction-- %08lx %08lx\n", &v, v.vt);
#endif
}

OCVariant::OCVariant(BSTR bstrVal) : next(NULL)
{
  VariantInit(&v);
  v.vt = VT_BSTR;
  v.bstrVal = bstrVal;
#ifdef DEBUG
  fprintf(stderr, "--construction-- %08lx %08lx\n", &v, v.vt);
#endif
}

OCVariant::OCVariant(string str) : next(NULL)
{
  VariantInit(&v);
  v.vt = VT_BSTR;
  v.bstrVal = SJIS2BSTR(str);
#ifdef DEBUG
  fprintf(stderr, "--construction-- %08lx %08lx\n", &v, v.vt);
#endif
}

OCVariant::~OCVariant()
{
  // bug ? comment (see old ole32core.cpp project)
  for(OCVariant *p = next, *q = NULL; p; p = q){
    q = p->next;
    delete p;
  }
  // The first node (== self) only be reversed.
  // 1, n, ..., 5, 4, 3, 2
#ifdef DEBUG
  fprintf(stderr, "--destruction-- %08lx %08lx\n", &v, v.vt);
#endif
  // bug ? comment (see old ole32core.cpp project)
  if((v.vt == VT_DISPATCH) && v.pdispVal){
    v.pdispVal->Release();
    VariantClear(&v);
    v.pdispVal = NULL;
  }
  if((v.vt == VT_BSTR) && v.bstrVal){
    ::SysFreeString(v.bstrVal);
    VariantClear(&v);
    v.bstrVal = NULL;
  }
}

OCVariant *OCVariant::push(OCVariant *p)
{
  // The first node (== self) only be reversed.
  // 1, n, ..., 5, 4, 3, 2
  p->next = next;
  next = p;
  return this;
}

unsigned int OCVariant::size()
{
  unsigned int c = 1; // for self
  for(OCVariant *p = next; p; p = p->next){
    c++;
  }
  return c;
}

void OCVariant::checkOLEresult(string msg)
{
  // bug ? comment (see old ole32core.cpp project)
  throw OLE32coreException(msg + "() failed\n");
}

// AutoWrap() - Automation helper function...
HRESULT OCVariant::AutoWrap(int autoType, VARIANT *pvResult,
  LPOLESTR ptName, OCVariant *argchain)
{
  // bug ? comment (see old ole32core.cpp project)
  // execute at the first time to safety free argchain
  // Allocate memory for arguments...
  unsigned int size = argchain ? argchain->size() : 0;
  VARIANT *pArgs = new VARIANT[size];
  OCVariant *p = argchain;
  for(unsigned int i = 0; p; i++, p = p->next){
    // bug ? comment (see old ole32core.cpp project)
    // will be reallocated BSTR whein using VariantCopy() (see by debugger)
    VariantInit(&pArgs[i]); // It will be free before copy.
    VariantCopy(&pArgs[i], &p->v);
  }
  if(argchain) delete argchain;
  // bug ? comment (see old ole32core.cpp project)
  // unexpected free original BSTR
  HRESULT hr = NULL;
  if(!v.pdispVal){
    checkOLEresult("Called with NULL IDispatch. AutoWrap");
    return hr;
  }
  // Convert down to ANSI (for error message only)
  char szName[256];
  WideCharToMultiByte(CP_ACP, 0,
    ptName, -1, szName, sizeof(szName), NULL, NULL);
  // Get DISPID for name passed...
  DISPID dispID;
  hr = v.pdispVal->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT,
    &dispID);
  if(FAILED(hr)){
    ostringstream oss;
    oss << hr << " [" << szName << "] ";
    oss << "IDispatch::GetIDsOfNames AutoWrap";
    checkOLEresult(oss.str());
    return hr;
  }
  // Build DISPPARAMS
  DISPPARAMS dp = { NULL, NULL, 0, 0 };
  dp.cArgs = size;
  dp.rgvarg = pArgs;
  // Handle special-case for property-puts!
  DISPID dispidNamed = DISPID_PROPERTYPUT;
  if(autoType & DISPATCH_PROPERTYPUT){
    dp.cNamedArgs = 1;
    dp.rgdispidNamedArgs = &dispidNamed;
  }
  // Make the call!
  hr = v.pdispVal->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT,
    autoType, &dp, pvResult, NULL, NULL);
  delete [] pArgs;
  if(FAILED(hr)){
    ostringstream oss;
    oss << hr << " [" << szName << "] = [" << dispID << "] ";
    oss << "IDispatch::Invoke AutoWrap";
    checkOLEresult(oss.str());
    return hr;
  }
  return hr;
}

OCVariant *OCVariant::getProp(LPOLESTR prop, OCVariant *argchain)
{
  OCVariant *r = new OCVariant();
  AutoWrap(DISPATCH_PROPERTYGET, &r->v, prop, argchain);
  return r;
}

OCVariant *OCVariant::putProp(LPOLESTR prop, OCVariant *argchain)
{
  AutoWrap(DISPATCH_PROPERTYPUT, NULL, prop, argchain);
  return this;
}

OCVariant *OCVariant::invoke(LPOLESTR method, OCVariant *argchain, bool re)
{
  if(!re){
    AutoWrap(DISPATCH_METHOD, NULL, method, argchain);
    return this;
  }else{
    OCVariant *r = new OCVariant();
    AutoWrap(DISPATCH_METHOD, &r->v, method, argchain);
    return r;
  }
}

void OLE32core::checkOLEresult(string msg)
{
  throw OLE32coreException(msg + "() failed\n");
}

OCVariant *OLE32core::connect(string locale, int visible)
{
  oldlocale = setlocale(LC_ALL, NULL);
  setlocale(LC_ALL, locale.c_str());
  CoInitialize(NULL); // Initialize COM for this thread...
  // Get CLSID for our server...
  if(FAILED(CLSIDFromProgID(L"Excel.Application", &clsid))){
    checkOLEresult("CLSIDFromProgID");
  }
  // Start server and get IDispatch... (app = CreateObject("..."))
  OCVariant *app = new OCVariant();
  app->v.vt = VT_DISPATCH;
  if(FAILED(CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER,
    IID_IDispatch, (void **)&app->v.pdispVal))){
    checkOLEresult("Excel is not registered ? CoCreateInstance");
  }
  if(visible){ // Make it visible (app.Visible = 1)
    app->putProp(L"Visible", new OCVariant((long)1));
  }
  return app;
}

void OLE32core::disconnect()
{
  CoUninitialize(); // Uninitialize COM for this thread...
#ifdef DEBUG
  setlocale(LC_ALL, "C");
#else
  setlocale(LC_ALL, oldlocale.c_str());
#endif
}
