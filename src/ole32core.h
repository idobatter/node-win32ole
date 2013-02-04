#ifndef __OLE32CORE_H__
#define __OLE32CORE_H__

#include <ole2.h>
#include <locale.h>

#include <vector>
#include <iomanip>
#include <iostream>
#include <sstream>
#include <string>

#include <cstdio>
#include <cstdlib>
#include <cstring>

namespace ole32core {

#ifdef DEBUG
#define DISPFUNCIN() do{std::cerr<<"-IN "<<__FUNCTION__<<std::endl;}while(0)
#define DISPFUNCOUT() do{std::cerr<<"-OUT "<<__FUNCTION__<<std::endl;}while(0)
#define DISPFUNCDAT(p, q, r) do{fprintf(stderr, p. q, r);}while(0)
#else
#define DISPFUNCIN() do{std::cerr<<"-IN "<<__FUNCTION__<<std::endl;}while(0)
#define DISPFUNCOUT() do{std::cerr<<"-OUT "<<__FUNCTION__<<std::endl;}while(0)
#define DISPFUNCDAT(p, q, r) do{fprintf(stderr, p, q, r);}while(0)
#endif

#define BASSERT(x) chkerr((BOOL)(x), __FILE__, __LINE__, __FUNCTION__, #x)
#define BVERIFY(x) BASSERT(x)
#if defined(_DEBUG) || defined(DEBUG)
#define DASSERT(x) BASSERT(x)
#define DVERIFY(x) DASSERT(x)
#else // RELEASE
#define DASSERT(x)
#define DVERIFY(x) (x)
#endif
extern BOOL chkerr(BOOL b, char *m, int n, char *f, char *e);
#define BEVERIFY(y, x) if(!BVERIFY(x)){ goto y; }
#define DEVERIFY(y, x) if(!DVERIFY(x)){ goto y; }

extern wchar_t *u8s2wcs(char *u8s); // UTF8 -> UCS2 (allocate wcs, must free)
extern char *wcs2mbs(wchar_t *wcs); // UCS2 -> locale (allocate mbs, must free)

// obsoleted functions

// (without free when use _malloca)
extern std::string MBCS2UTF8(std::string mbs);
// (without free when use _malloca)
extern std::string UTF82MBCS(std::string utf8);

// locale mbs -> Unicode (allocate wcs, must free)
extern WCHAR *MBCS2WCS(std::string mbs);
// Unicode -> locale mbs (not free wcs, without free when use _malloca)
extern std::string WCS2MBCS(WCHAR *wbuf);

// (allocate bstr, must free)
extern BSTR MBCS2BSTR(std::string str);
// (not free bstr)
extern std::string BSTR2MBCS(BSTR bstr);

class OLE32coreException {
protected:
  std::string rmsg;
public:
  OLE32coreException(std::string rm) : rmsg(rm) {}
  virtual ~OLE32coreException() {}
  std::string errorMessage(char *m=NULL);
};

class OCVariant {
public:
  VARIANT v;
  OCVariant *next;
public:
  OCVariant(); // result
  OCVariant(const OCVariant &s); // copy
  OCVariant(long lVal); // VT_I4
  OCVariant(double dblVal); // VT_R8
  OCVariant(BSTR bstrVal); // VT_BSTR (previous allocated)
  OCVariant(std::string str); // allocate and convert to VT_BSTR
  virtual ~OCVariant();
  OCVariant *push(OCVariant *p); // push to chain top
  unsigned int size(); // length of chain (count next and self)
  void checkOLEresult(std::string msg);
protected:
  HRESULT AutoWrap(int autoType, VARIANT *pvResult,
    LPOLESTR ptName, OCVariant *argchain=NULL);
public:
  OCVariant *getProp(LPOLESTR prop, OCVariant *argchain=NULL);
  OCVariant *putProp(LPOLESTR prop, OCVariant *argchain=NULL);
  OCVariant *invoke(LPOLESTR method, OCVariant *argchain=NULL, bool re=false);
};

class OLE32core {
protected:
  bool finalized;
  std::string oldlocale;
public:
  OLE32core() : finalized(true) {}
  virtual ~OLE32core() { if(!finalized) disconnect(); }
  void checkOLEresult(std::string msg);
  bool connect(std::string locale);
  bool disconnect(void);
};

} // namespace ole32core

#endif // __OLE32CORE_H__
