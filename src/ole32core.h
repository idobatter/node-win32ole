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
#include <ctime>

namespace ole32core {

#define BDISPFUNCIN() do{std::cerr<<"-IN "<<__FUNCTION__<<std::endl;}while(0)
#define BDISPFUNCOUT() do{std::cerr<<"-OUT "<<__FUNCTION__<<std::endl;}while(0)
#define BDISPFUNCDAT(f, a, t) do{fprintf(stderr, (f), (a), (t));}while(0)
#if defined(_DEBUG) || defined(DEBUG)
#define DISPFUNCIN() BDISPFUNCIN()
#define DISPFUNCOUT() BDISPFUNCOUT()
#define DISPFUNCDAT(f, a, t) BDISPFUNCDAT((f), (a), (t))
#else // RELEASE
#define DISPFUNCIN() // BDISPFUNCIN()
#define DISPFUNCOUT() // BDISPFUNCOUT()
#define DISPFUNCDAT(f, a, t) // BDISPFUNCDAT((f), (a), (t))
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

extern std::string to_s(int num);
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
  OCVariant(bool c_boolVal); // VT_BOOL
  OCVariant(long lVal); // VT_I4
  OCVariant(double dblVal); // VT_R8
  OCVariant(double date, bool isdate); // VT_DATE
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
