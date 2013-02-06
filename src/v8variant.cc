/*
  v8variant.cc
*/

#include "v8variant.h"

using namespace v8;
using namespace ole32core;

namespace node_win32ole {

#define CHECK_OCV(s, ocv) do{ \
    if(!ocv) \
      return ThrowException(Exception::TypeError( \
        String::New( s " can't access to V8Variant (null OCVariant)"))); \
  }while(0)

#define CHECK_OLE_ARGS(s, args, n, av0, av1) do{ \
    if(args.Length() < n) \
      return ThrowException(Exception::TypeError( \
        String::New( s " takes exactly " #n " argument(s)"))); \
    if(!args[0]->IsString()) \
      return ThrowException(Exception::TypeError( \
        String::New( s " the first argument is not a Symbol"))); \
    if(n == 1) \
      if(args.Length() >= 2) \
        if(!args[1]->IsArray()) \
          return ThrowException(Exception::TypeError( \
            String::New( s " the second argument is not an Array"))); \
        else av1 = args[1]; /* Array */ \
      else av1 = Array::New(0); /* change none to Array[] */ \
    else av1 = args[1]; /* may not be Array */ \
    av0 = args[0]; \
  }while(0)

Persistent<FunctionTemplate> V8Variant::clazz;

void V8Variant::Init(Handle<Object> target)
{
  HandleScope scope;
  Local<FunctionTemplate> t = FunctionTemplate::New(New);
  clazz = Persistent<FunctionTemplate>::New(t);
  clazz->InstanceTemplate()->SetInternalFieldCount(1);
  clazz->SetClassName(String::NewSymbol("V8Variant"));
  NODE_SET_PROTOTYPE_METHOD(clazz, "toBoolean", OLEBoolean);
  NODE_SET_PROTOTYPE_METHOD(clazz, "toInt32", OLEInt32);
  NODE_SET_PROTOTYPE_METHOD(clazz, "toNumber", OLENumber);
  NODE_SET_PROTOTYPE_METHOD(clazz, "toUtf8", OLEUtf8);
//  NODE_SET_PROTOTYPE_METHOD(clazz, "New", New);
  NODE_SET_PROTOTYPE_METHOD(clazz, "get", OLEGet);
  NODE_SET_PROTOTYPE_METHOD(clazz, "set", OLESet);
  NODE_SET_PROTOTYPE_METHOD(clazz, "call", OLECall);
  NODE_SET_PROTOTYPE_METHOD(clazz, "Finalize", Finalize);
  target->Set(String::NewSymbol("V8Variant"), clazz->GetFunction());
}

std::string V8Variant::CreateStdStringMBCSfromUTF8(Handle<Value> v)
{
  String::Utf8Value u8s(v);
  wchar_t * wcs = u8s2wcs(*u8s);
  if(!wcs){
    std::cerr << "[Can't allocate string (wcs)]" << std::endl;
    return std::string("'!ERROR'");
  }
  char *mbs = wcs2mbs(wcs);
  if(!mbs){
    free(wcs);
    std::cerr << "[Can't allocate string (mbs)]" << std::endl;
    return std::string("'!ERROR'");
  }
  std::string s(mbs);
  free(mbs);
  free(wcs);
  return s;
}

OCVariant *V8Variant::CreateOCVariant(Handle<Value> v)
{
  BEVERIFY(done, !v.IsEmpty());
  BEVERIFY(done, !v->IsUndefined());
  BEVERIFY(done, !v->IsNull());
  BEVERIFY(done, !v->IsExternal());
  BEVERIFY(done, !v->IsNativeError());
  BEVERIFY(done, !v->IsFunction());
// VT_USERDEFINED VT_VARIANT VT_BYREF VT_ARRAY more...
  if(v->IsBoolean()){
    return new OCVariant((bool)(v->BooleanValue() ? !0 : 0));
  }else if(v->IsArray()){
// VT_BYREF VT_ARRAY VT_SAFEARRAY
    std::cerr << "[Array (not implemented now)]" << std::endl; return NULL;
  }else if(v->IsInt32()){
    return new OCVariant((long)v->Int32Value());
  }else if(v->IsNumber()){
    std::cerr << "[Number (VT_R8 or VT_I8 bug?)]" << std::endl;
// if(v->ToInteger()) =64 is failed ? double : OCVariant((longlong)VT_I8)
    return new OCVariant((double)v->NumberValue()); // double
  }else if(v->IsNumberObject()){
    std::cerr << "[NumberObject (VT_R8 or VT_I8 bug?)]" << std::endl;
// if(v->ToInteger()) =64 is failed ? double : OCVariant((longlong)VT_I8)
    return new OCVariant((double)v->NumberValue()); // double
  }else if(v->IsDate()){
    std::cerr << "[Date (bug?)]" << std::endl;
    return new OCVariant(CreateStdStringMBCSfromUTF8(v->ToDetailString()));
  }else if(v->IsRegExp()){
    std::cerr << "[RegExp (bug?)]" << std::endl;
    return new OCVariant(CreateStdStringMBCSfromUTF8(v->ToDetailString()));
  }else if(v->IsString()){
    return new OCVariant(CreateStdStringMBCSfromUTF8(v));
  }else if(v->IsStringObject()){
    std::cerr << "[StringObject (bug?)]" << std::endl;
    return new OCVariant(CreateStdStringMBCSfromUTF8(v));
  }else if(v->IsObject()){
    std::cerr << "[Object (test)]" << std::endl;
    OCVariant *ocv = castedInternalField<OCVariant>(v->ToObject());
    if(!ocv){
      std::cerr << "[Object may not be valid (null OCVariant)]" << std::endl;
      return NULL;
    }
    return new OCVariant(*ocv);
  }else{
    std::cerr << "[unknown type (not implemented now)]" << std::endl;
  }
done:
  return NULL;
}

Handle<Value> V8Variant::OLEBoolean(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  CHECK_OCV("OLEBoolean", ocv);
  if(ocv->v.vt != VT_BOOL)
    return ThrowException(Exception::TypeError(
      String::New("OLEBoolean source type OCVariant is not VT_BOOL")));
  bool c_boolVal = ocv->v.boolVal == VARIANT_FALSE ? 0 : !0;
  DISPFUNCOUT();
  return scope.Close(Boolean::New(c_boolVal));
}

Handle<Value> V8Variant::OLEInt32(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  CHECK_OCV("OLEInt32", ocv);
#if(0)
  if(ocv->v.vt != VT_I4 && ocv->v.vt != VT_INT)
    return ThrowException(Exception::TypeError(
      String::New("OLEInt32 source type OCVariant is not VT_I4 nor VT_INT")));
  if(ocv->v.vt != VT_UI4 && ocv->v.vt != VT_UINT)
    return ThrowException(Exception::TypeError(
      String::New("OLEUInt32 source type OCVariant is not VT_UI4 nor VT_UINT")));
  if(ocv->v.vt != VT_I4 && ocv->v.vt != VT_INT)
    std::cerr << "[Int32 (not VT_I4 nor VT_INT bug?)]" << ocv->v.vt << std::endl;
  if(ocv->v.vt != VT_UI4 && ocv->v.vt != VT_UINT)
    std::cerr << "[UInt32 (not VT_UI4 nor VT_UINT bug?)]" << ocv->v.vt << std::endl;
  if(ocv->v.vt != VT_I8 && ocv->v.vt != VT_UI8)
    std::cerr << "[(U)Int64 (not VT_I8 nor VT_UI8 bug?)]" << ocv->v.vt << std::endl;
#else
  if(ocv->v.vt != VT_I4 && ocv->v.vt != VT_INT
  && ocv->v.vt != VT_UI4 && ocv->v.vt != VT_UINT)
    return ThrowException(Exception::TypeError(
      String::New("OLEInt32 source type OCVariant is not VT_I4 nor VT_INT nor VT_UI4 nor VT_UINT")));
#endif
  DISPFUNCOUT();
  return scope.Close(Int32::New(ocv->v.lVal));
}

Handle<Value> V8Variant::OLENumber(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  CHECK_OCV("OLENumber", ocv);
  if(ocv->v.vt != VT_R8)
    return ThrowException(Exception::TypeError(
      String::New("OLENumber source type OCVariant is not VT_R8")));
  DISPFUNCOUT();
  return scope.Close(Number::New(ocv->v.dblVal));
}

Handle<Value> V8Variant::OLEUtf8(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  CHECK_OCV("OLEUtf8", ocv);
  if(ocv->v.vt != VT_BSTR)
    return ThrowException(Exception::TypeError(
      String::New("OLEUtf8 source type OCVariant is not VT_BSTR")));
  Handle<Value> result;
  if(!ocv->v.bstrVal) result = Undefined(); // Null();
  else result = String::New(MBCS2UTF8(BSTR2MBCS(ocv->v.bstrVal)).c_str());
  DISPFUNCOUT();
  return scope.Close(result);
}

Handle<Object> V8Variant::CreateUndefined(void)
{
  HandleScope scope;
  DISPFUNCIN();
#if(0)
  const size_t argc = 1;
  Handle<Value> argv[argc] = {args[0]};
  // Class {}
  //   static Persistent<Function> V8Variant::constructor_func;
  // Init()
  //   constructor_func = Persistent<Function>::New(clazz->GetFunction());
  Local<Object> instance = constructor->NewInstance(argc, argv);
#else
  Local<Object> instance = clazz->GetFunction()->NewInstance(0, NULL);
#endif
  DISPFUNCOUT();
  return scope.Close(instance);
}

Handle<Value> V8Variant::New(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
  if(!args.IsConstructCall())
    return ThrowException(Exception::TypeError(
      String::New("Use the new operator to create new V8Variant objects")));
  OCVariant *ocv = new OCVariant();
  CHECK_OCV("New", ocv);
  Local<Object> thisObject = args.This();
  thisObject->SetInternalField(0, External::New(ocv));
  Persistent<Object> objectDisposer = Persistent<Object>::New(thisObject);
  objectDisposer.MakeWeak(ocv, Dispose);
  DISPFUNCOUT();
  return args.This();
}

Handle<Value> V8Variant::OLEGet(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  CHECK_OCV("OLEGet", ocv);
  Handle<Value> av0, av1;
  CHECK_OLE_ARGS("OLEGet", args, 1, av0, av1);
  OCVariant *argchain = NULL;
  Array *a = Array::Cast(*av1);
  for(size_t i = 0; i < a->Length(); ++i){
    char num[256];
    sprintf(num, "%d", a->Length() - 1 - i); // *** check length ***
    OCVariant *o = CreateOCVariant(a->Get(String::NewSymbol(num)));
    if(!o)
      return ThrowException(Exception::TypeError(
        String::New("OLEGet can't access to argument i (null OCVariant)")));
    if(!i) argchain = o;
    else argchain->push(o);
  }
  Handle<Object> vResult = V8Variant::CreateUndefined();
  String::Utf8Value u8s(av0);
  wchar_t *wcs = u8s2wcs(*u8s);
  if(!wcs && argchain) delete argchain;
  BEVERIFY(done, wcs);
  try{
    OCVariant *rv = ocv->getProp(wcs, argchain); // argchain will be deleted
    if(rv){
      OCVariant *o = castedInternalField<OCVariant>(vResult);
      CHECK_OCV("OLEGet(result)", o);
      *o = *rv; // copy and don't delete rv
    }
  }catch(OLE32coreException e){ std::cerr << e.errorMessage(*u8s); goto done;
  }catch(char *e){ std::cerr << e << *u8s << std::endl; goto done;
  }
  free(wcs); // *** it may leak when error ***
  DISPFUNCOUT();
  return scope.Close(vResult);
done:
  DISPFUNCOUT();
  return ThrowException(Exception::TypeError(String::New("OLEGet failed")));
}

Handle<Value> V8Variant::OLESet(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  CHECK_OCV("OLESet", ocv);
  Handle<Value> av0, av1;
  CHECK_OLE_ARGS("OLESet", args, 2, av0, av1);
  OCVariant *a1 = CreateOCVariant(av1); // will be deleted automatically
  if(!a1)
    return ThrowException(Exception::TypeError(
      String::New("the second argument is not valid (null OCVariant)")));
  bool result = false;
  String::Utf8Value u8s(av0);
  wchar_t *wcs = u8s2wcs(*u8s);
  BEVERIFY(done, wcs);
  try{
    ocv->putProp(wcs, a1);
  }catch(OLE32coreException e){ std::cerr << e.errorMessage(*u8s); goto done;
  }catch(char *e){ std::cerr << e << *u8s << std::endl; goto done;
  }
  free(wcs);
  result = true;
  DISPFUNCOUT();
  return scope.Close(Boolean::New(result));
done:
  DISPFUNCOUT();
  return ThrowException(Exception::TypeError(String::New("OLESet failed")));
}

Handle<Value> V8Variant::OLECall(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  CHECK_OCV("OLECall", ocv);
  Handle<Value> av0, av1;
  CHECK_OLE_ARGS("OLECall", args, 1, av0, av1);
  OCVariant *argchain = NULL;
  Array *a = Array::Cast(*av1);
  for(size_t i = 0; i < a->Length(); ++i){
    char num[256];
    sprintf(num, "%d", a->Length() - 1 - i); // *** check length ***
    OCVariant *o = CreateOCVariant(a->Get(String::NewSymbol(num)));
    if(!o)
      return ThrowException(Exception::TypeError(
        String::New("OLECall can't access to argument i (null OCVariant)")));
    if(!i) argchain = o;
    else argchain->push(o);
  }
  Handle<Object> vResult = V8Variant::CreateUndefined();
  String::Utf8Value u8s(av0);
  wchar_t *wcs = u8s2wcs(*u8s);
  if(!wcs && argchain) delete argchain;
  BEVERIFY(done, wcs);
  try{
    OCVariant *rv = ocv->invoke(wcs, argchain, true); // argchain will be deleted
    if(rv){
      OCVariant *o = castedInternalField<OCVariant>(vResult);
      CHECK_OCV("OLECall(result)", o);
      *o = *rv; // copy and don't delete rv
    }
  }catch(OLE32coreException e){ std::cerr << e.errorMessage(*u8s); goto done;
  }catch(char *e){ std::cerr << e << *u8s << std::endl; goto done;
  }
  free(wcs); // *** it may leak when error ***
  DISPFUNCOUT();
  return scope.Close(vResult);
done:
  DISPFUNCOUT();
  return ThrowException(Exception::TypeError(String::New("OLECall failed")));
}

Handle<Value> V8Variant::Finalize(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
#if(0)
  std::cerr << __FUNCTION__ << " Finalizer is called\a" << std::endl;
#endif
  Local<Object> thisObject = args.This();
#if(1) // now GC will call Disposer automatically
  OCVariant *ocv = castedInternalField<OCVariant>(thisObject);
  if(ocv) delete ocv;
#endif
  thisObject->SetInternalField(0, External::New(NULL));
#if(0) // to force GC (shold not use at here ?)
  // v8::internal::Heap::CollectAllGarbage(); // obsolete
  while(!v8::V8::IdleNotification());
#endif
  DISPFUNCOUT();
  return args.This();
}

void V8Variant::Dispose(Persistent<Value> handle, void *param)
{
  DISPFUNCIN();
#if(1)
  std::cerr << __FUNCTION__ << " Disposer is called\a" << std::endl;
#endif
  OCVariant *p = castedInternalField<OCVariant>(handle->ToObject());
  if(!p){
    std::cerr << __FUNCTION__;
    std::cerr << "InternalField has been already deleted" << std::endl;
  }
//  else{
    OCVariant *ocv = static_cast<OCVariant *>(param);
    if(ocv) delete ocv;
//  }
  handle->ToObject()->SetInternalField(0, External::New(NULL));
  handle.Dispose();
  DISPFUNCOUT();
}

void V8Variant::Finalize()
{
  assert(!finalized);
  finalized = true;
}

} // namespace node_win32ole
