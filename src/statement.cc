/*
  statement.cc
*/

#include "statement.h"
#include "v8variant.h"

using namespace v8;
using namespace ole32core;

namespace node_win32ole {

Persistent<FunctionTemplate> Statement::clazz;

void Statement::Init(Handle<Object> target)
{
  HandleScope scope;
  Local<FunctionTemplate> t = FunctionTemplate::New(New);
  clazz = Persistent<FunctionTemplate>::New(t);
  clazz->InstanceTemplate()->SetInternalFieldCount(1);
  clazz->SetClassName(String::NewSymbol("Statement"));
//  NODE_SET_PROTOTYPE_METHOD(clazz, "New", New);
  NODE_SET_PROTOTYPE_METHOD(clazz, "Dispatch", Dispatch);
  NODE_SET_PROTOTYPE_METHOD(clazz, "Finalize", Finalize);
  target->Set(String::NewSymbol("Statement"), clazz->GetFunction());
}

Handle<Value> Statement::New(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
  if(!args.IsConstructCall())
    return ThrowException(Exception::TypeError(
      String::New("Use the new operator to create new Statement objects")));
  OLE32core *oc = new OLE32core();
  if(!oc)
    return ThrowException(Exception::TypeError(
      String::New("Can't create new Statement object (null OLE32core)")));
  Local<Object> thisObject = args.This();
  thisObject->SetInternalField(0, External::New(oc));
  Persistent<Object> objectDisposer = Persistent<Object>::New(thisObject);
  objectDisposer.MakeWeak(oc, Dispose);
  DISPFUNCOUT();
  return args.This();
}

Handle<Value> Statement::Dispatch(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
  BEVERIFY(done, args.Length() >= 2);
  BEVERIFY(done, args[0]->IsString());
  BEVERIFY(done, args[1]->IsString());
  char cstr_locale[256];
  {
    String::Utf8Value u8s_locale(args[1]);
    strncpy(cstr_locale, *u8s_locale, sizeof(cstr_locale));
  }
  wchar_t *wcs;
  {
    String::Utf8Value u8s(args[0]); // must create here
    wcs = u8s2wcs(*u8s);
  }
  BEVERIFY(done, wcs);
#ifdef DEBUG
  char *mbs = wcs2mbs(wcs);
  if(!mbs) free(wcs);
  BEVERIFY(done, mbs);
  fprintf(stderr, "ProgID: %s\n", mbs);
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
  OLE32core *oc = castedInternalField<OLE32core>(args.This());
  if(!oc)
    return ThrowException(Exception::TypeError(
      String::New("Can't access to Statement object (null OLE32core)")));
  BEVERIFY(done, oc->connect(cstr_locale));
  Handle<Object> vApp = V8Variant::CreateUndefined();
  BEVERIFY(done, !vApp.IsEmpty());
  BEVERIFY(done, !vApp->IsUndefined());
  BEVERIFY(done, vApp->IsObject());
  OCVariant *app = castedInternalField<OCVariant>(vApp);
  if(!app)
    return ThrowException(Exception::TypeError(
      String::New("Can't access to V8Variant object (null OCVariant)")));
  app->v.vt = VT_DISPATCH;
  // When 'CoInitialize(NULL)' is not called first (and on the same instance),
  // next functions will return many errors.
  // (old style) GetActiveObject() returns 0x000036b7
  //   The requested lookup key was not found in any active activation context.
  // (OLE2) CoCreateInstance() returns 0x000003f0
  //   An attempt was made to reference a token that does not exist.
#ifdef DEBUG // obsolete (it needs that OLE target has been already executed)
  IUnknown *pUnk;
  hr = GetActiveObject(clsid, NULL, (IUnknown **)&pUnk);
  BEVERIFY(done, !FAILED(hr));
  hr = pUnk->QueryInterface(IID_IDispatch, (void **)&app->v.pdispVal);
  pUnk->Release();
#else
  // C -> C++ changes types (&clsid -> clsid, &IID_IDispatch -> IID_IDispatch)
  hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER|CLSCTX_LOCAL_SERVER,
    IID_IDispatch, (void **)&app->v.pdispVal);
#endif
  BEVERIFY(done, !FAILED(hr));
  DISPFUNCOUT();
  return scope.Close(vApp);
done:
  DISPFUNCOUT();
  return ThrowException(Exception::TypeError(String::New("Dispatch failed")));
}

Handle<Value> Statement::Finalize(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
#if(0)
  std::cerr << __FUNCTION__ << " Finalizer is called\a" << std::endl;
#endif
  Local<Object> thisObject = args.This();
#if(1) // now GC will call Disposer automatically
  OLE32core *oc = castedInternalField<OLE32core>(thisObject);
  if(oc){ oc->disconnect(); delete oc; }
#endif
  thisObject->SetInternalField(0, External::New(NULL));
#if(0) // to force GC (shold not use at here ?)
  // v8::internal::Heap::CollectAllGarbage(); // obsolete
  while(!v8::V8::IdleNotification());
#endif
  DISPFUNCOUT();
  return args.This();
}

void Statement::Dispose(Persistent<Value> handle, void *param)
{
  DISPFUNCIN();
#if(1)
  std::cerr << __FUNCTION__ << " Disposer is called\a" << std::endl;
#endif
  OLE32core *p = castedInternalField<OLE32core>(handle->ToObject());
  if(!p){
    std::cerr << __FUNCTION__;
    std::cerr << " InternalField has been already deleted" << std::endl;
  }
//  else{
    OLE32core *oc = static_cast<OLE32core *>(param);
    if(oc){ oc->disconnect(); delete oc; }
//  }
  handle->ToObject()->SetInternalField(0, External::New(NULL));
  handle.Dispose();
  DISPFUNCOUT();
}

void Statement::Finalize()
{
  assert(!finalized);
  finalized = true;
}

} // namespace node_win32ole
