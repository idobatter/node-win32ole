/*
  node_win32ole.cc
  $ cp ~/.node-gyp/0.8.18/ia32/node.lib to ~/.node-gyp/0.8.18/node.lib
  $ node-gyp configure
  $ node-gyp build
  $ node test/init_win32ole.test.js
*/

#include <node_buffer.h>
#include <node.h>
#include <v8.h>

#include "node_win32ole.h"

using namespace v8;

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

Persistent<Object> module_target;
Persistent<FunctionTemplate> Statement::constructor_template;
VARIANT Statement::vDisp = {0};

void Statement::Init(Handle<Object> target)
{
  HandleScope scope;
  Local<FunctionTemplate> t = FunctionTemplate::New(New);
  constructor_template = Persistent<FunctionTemplate>::New(t);
  constructor_template->InstanceTemplate()->SetInternalFieldCount(1);
  constructor_template->SetClassName(String::NewSymbol("Statement"));
  NODE_SET_PROTOTYPE_METHOD(constructor_template, "Dispatch", Dispatch);
  NODE_SET_PROTOTYPE_METHOD(constructor_template, "Finalize", Finalize);
  target->Set(String::NewSymbol("Statement"),
    constructor_template->GetFunction());
}

Handle<Value> Statement::New(const Arguments& args)
{
  HandleScope scope;
  if(!args.IsConstructCall())
    return ThrowException(Exception::TypeError(
      String::New("Use the new operator to create new Statement objects")));
  return args.This();
}

Handle<Value> Statement::Dispatch(const Arguments& args)
{
  HandleScope scope;
  boolean result = false;
  if(!args[0]->IsString()){
    fprintf(stderr, "error: args[0] is not a string");
    return scope.Close(Boolean::New(result));
  }
  String::Utf8Value s(args[0]);
  wchar_t *wcs = u8s2wcs(*s);
  BEVERIFY(done, wcs);
#ifdef DEBUG
  char *mbs = wcs2mbs(wcs);
  if(!mbs) free(wcs);
  BEVERIFY(done, mbs);
  fprintf(stderr, "ProgID: %s\n", mbs); // Excel.Application
  free(mbs);
#endif
  CLSID clsid;
  HRESULT hr = CLSIDFromProgID(wcs, &clsid);
  free(wcs);
  BEVERIFY(done, !FAILED(hr));
#ifdef DEBUG
  fprintf(stderr, "clsid:"); // 00024500-0000-0000-c000-000000000046 (Excel) ok
  for(int i = 0; i < sizeof(CLSID); ++i)
    fprintf(stderr, " %02x", ((unsigned char *)&clsid)[i]);
  fprintf(stderr, "\n");
#endif
  // When this function is not called first (and on the same instance),
  // next functions will return many errors.
  CoInitialize(NULL);
  // (old style) GetActiveObject() returns 0x000036b7
  //   The requested lookup key was not found in any active activation context.
  // (OLE2) CoCreateInstance() returns 0x000003f0
  //   An attempt was made to reference a token that does not exist.
#ifdef DEBUG // old style (needs pre executed)
  IUnknown *pUnk;
  hr = GetActiveObject(clsid, NULL, (IUnknown **)&pUnk);
  BEVERIFY(done, !FAILED(hr));
  hr = pUnk->QueryInterface(IID_IDispatch, (void **)&vDisp.pdispVal);
  BEVERIFY(done, !FAILED(hr));
  pUnk->Release();
#else
  // C -> C++ changes types (&clsid -> clsid, &IID_IDispatch -> IID_IDispatch)
  vDisp.vt = VT_DISPATCH;
  hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER|CLSCTX_LOCAL_SERVER,
    IID_IDispatch, (void **)&vDisp.pdispVal);
  BEVERIFY(done, !FAILED(hr));
#endif
  result = true;
done:
  return scope.Close(Boolean::New(result));
}

Handle<Value> Statement::Finalize(const Arguments& args)
{
  HandleScope scope;
  if(vDisp.vt == VT_DISPATCH && vDisp.pdispVal)
    VariantClear(&vDisp); // vDisp.pdispVal->Release();
  CoUninitialize();
  return args.This();
}

void Statement::Finalize()
{
  assert(!finalized);
  finalized = true;
}

Handle<Value> Method_version(const Arguments& args)
{
  HandleScope scope;
  if(args.Length() >= 0){
    //
  }
#ifdef DEBUG
  node::node_module_struct *nms = node::get_builtin_module("Array");
  if(!nms) return scope.Close(String::NewSymbol("not found"));
  return scope.Close(String::New(nms->modname));
#else
  return scope.Close(module_target->Get(String::NewSymbol("VERSION")));
#endif
}

Handle<Value> Method_print(const Arguments& args)
{
  HandleScope scope;
  if(args.Length() >= 1){
    String::Utf8Value s(args[0]);
    char *cs = *s;
    printf("%s\n", cs);
  }
  return scope.Close(Boolean::New(true));
}

void init(Handle<Object> target)
{
  module_target = Persistent<Object>::New(target);
  Statement::Init(target);
  target->Set(String::NewSymbol("VERSION"),
    String::New("0.0.0 (will be set later)"),
    static_cast<PropertyAttribute>(DontDelete));
  target->Set(String::NewSymbol("SOURCE_TIMESTAMP"),
    String::NewSymbol(__FILE__ " " __DATE__ " " __TIME__),
    static_cast<PropertyAttribute>(ReadOnly | DontDelete));
  target->Set(String::NewSymbol("version"),
    FunctionTemplate::New(Method_version)->GetFunction());
  target->Set(String::NewSymbol("print"),
    FunctionTemplate::New(Method_print)->GetFunction());
  target->Set(String::NewSymbol("gettimeofday"),
    FunctionTemplate::New(Method_gettimeofday)->GetFunction());
}

NODE_MODULE(node_win32ole, init)
