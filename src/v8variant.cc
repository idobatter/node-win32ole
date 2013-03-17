/*
  v8variant.cc
*/

#include "v8variant.h"

using namespace v8;
using namespace ole32core;

namespace node_win32ole {

#define CHECK_OLE_ARGS(args, n, av0, av1) do{ \
    if(args.Length() < n) \
      return ThrowException(Exception::TypeError( \
        String::New(__FUNCTION__" takes exactly " #n " argument(s)"))); \
    if(!args[0]->IsString()) \
      return ThrowException(Exception::TypeError( \
        String::New(__FUNCTION__" the first argument is not a Symbol"))); \
    if(n == 1) \
      if(args.Length() >= 2) \
        if(!args[1]->IsArray()) \
          return ThrowException(Exception::TypeError(String::New( \
            __FUNCTION__" the second argument is not an Array"))); \
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
  clazz->InstanceTemplate()->SetInternalFieldCount(2);
  clazz->SetClassName(String::NewSymbol("V8Variant"));
  NODE_SET_PROTOTYPE_METHOD(clazz, "isA", OLEIsA);
  NODE_SET_PROTOTYPE_METHOD(clazz, "vtName", OLEVTName);
  NODE_SET_PROTOTYPE_METHOD(clazz, "toBoolean", OLEBoolean);
  NODE_SET_PROTOTYPE_METHOD(clazz, "toInt32", OLEInt32);
  NODE_SET_PROTOTYPE_METHOD(clazz, "toInt64", OLEInt64);
  NODE_SET_PROTOTYPE_METHOD(clazz, "toNumber", OLENumber);
  NODE_SET_PROTOTYPE_METHOD(clazz, "toDate", OLEDate);
  NODE_SET_PROTOTYPE_METHOD(clazz, "toUtf8", OLEUtf8);
  NODE_SET_PROTOTYPE_METHOD(clazz, "toValue", OLEValue);
//  NODE_SET_PROTOTYPE_METHOD(clazz, "New", New);
  NODE_SET_PROTOTYPE_METHOD(clazz, "call", OLECall);
  NODE_SET_PROTOTYPE_METHOD(clazz, "get", OLEGet);
  NODE_SET_PROTOTYPE_METHOD(clazz, "set", OLESet);
/*
 In ParseUnaryExpression() < v8/src/parser.cc >
 v8::Object::ToBoolean() is called directly for unary operator '!'
 instead of v8::Object::valueOf()
 so NamedPropertyHandler will not be called
 Local<Boolean> ToBoolean(); // How to fake ? override v8::Value::ToBoolean
*/
  Local<ObjectTemplate> instancetpl = clazz->InstanceTemplate();
  instancetpl->SetCallAsFunctionHandler(OLECallComplete);
  instancetpl->SetNamedPropertyHandler(OLEGetAttr, OLESetAttr);
//  instancetpl->SetIndexedPropertyHandler(OLEGetIdxAttr, OLESetIdxAttr);
  NODE_SET_PROTOTYPE_METHOD(clazz, "Finalize", Finalize);
  target->Set(String::NewSymbol("V8Variant"), clazz->GetFunction());
}

std::string V8Variant::CreateStdStringMBCSfromUTF8(Handle<Value> v)
{
  String::Utf8Value u8s(v);
  wchar_t * wcs = u8s2wcs(*u8s);
  if(!wcs){
    std::cerr << "[Can't allocate string (wcs)]" << std::endl;
    std::cerr.flush();
    return std::string("'!ERROR'");
  }
  char *mbs = wcs2mbs(wcs);
  if(!mbs){
    free(wcs);
    std::cerr << "[Can't allocate string (mbs)]" << std::endl;
    std::cerr.flush();
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
    std::cerr.flush();
  }else if(v->IsInt32()){
    return new OCVariant((long)v->Int32Value());
#if(0) // may not be supported node.js / v8
  }else if(v->IsInt64()){
    return new OCVariant((long long)v->Int64Value());
#endif
  }else if(v->IsNumber()){
    std::cerr << "[Number (VT_R8 or VT_I8 bug?)]" << std::endl;
    std::cerr.flush();
// if(v->ToInteger()) =64 is failed ? double : OCVariant((longlong)VT_I8)
    return new OCVariant((double)v->NumberValue()); // double
  }else if(v->IsNumberObject()){
    std::cerr << "[NumberObject (VT_R8 or VT_I8 bug?)]" << std::endl;
    std::cerr.flush();
// if(v->ToInteger()) =64 is failed ? double : OCVariant((longlong)VT_I8)
    return new OCVariant((double)v->NumberValue()); // double
  }else if(v->IsDate()){
    double d = v->NumberValue();
    time_t sec = (time_t)(d / 1000.0);
    int msec = (int)(d - sec * 1000.0);
    struct tm *t = localtime(&sec); // *** must check locale ***
    if(!t){
      std::cerr << "[Date may not be valid]" << std::endl;
      std::cerr.flush();
      return NULL;
    }
    SYSTEMTIME syst;
    syst.wYear = t->tm_year + 1900;
    syst.wMonth = t->tm_mon + 1;
    syst.wDay = t->tm_mday;
    syst.wHour = t->tm_hour;
    syst.wMinute = t->tm_min;
    syst.wSecond = t->tm_sec;
    syst.wMilliseconds = msec;
    SystemTimeToVariantTime(&syst, &d);
    return new OCVariant(d, true); // date
  }else if(v->IsRegExp()){
    std::cerr << "[RegExp (bug?)]" << std::endl;
    std::cerr.flush();
    return new OCVariant(CreateStdStringMBCSfromUTF8(v->ToDetailString()));
  }else if(v->IsString()){
    return new OCVariant(CreateStdStringMBCSfromUTF8(v));
  }else if(v->IsStringObject()){
    std::cerr << "[StringObject (bug?)]" << std::endl;
    std::cerr.flush();
    return new OCVariant(CreateStdStringMBCSfromUTF8(v));
  }else if(v->IsObject()){
#if(0)
    std::cerr << "[Object (test)]" << std::endl;
    std::cerr.flush();
#endif
    OCVariant *ocv = castedInternalField<OCVariant>(v->ToObject());
    if(!ocv){
      std::cerr << "[Object may not be valid (null OCVariant)]" << std::endl;
      std::cerr.flush();
      return NULL;
    }
    // std::cerr << ocv->v.vt;
    return new OCVariant(*ocv);
  }else{
    std::cerr << "[unknown type (not implemented now)]" << std::endl;
    std::cerr.flush();
  }
done:
  return NULL;
}

Handle<Value> V8Variant::OLEIsA(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  CHECK_OCV(ocv);
  DISPFUNCOUT();
  return scope.Close(Int32::New(ocv->v.vt));
}

Handle<Value> V8Variant::OLEVTName(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  CHECK_OCV(ocv);
  Array *a = Array::Cast(*GET_PROP(module_target, "vt_names"));
  DISPFUNCOUT();
  return scope.Close(ARRAY_AT(a, ocv->v.vt));
}

Handle<Value> V8Variant::OLEBoolean(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  CHECK_OCV(ocv);
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
  CHECK_OCV(ocv);
  if(ocv->v.vt != VT_I4 && ocv->v.vt != VT_INT
  && ocv->v.vt != VT_UI4 && ocv->v.vt != VT_UINT)
    return ThrowException(Exception::TypeError(
      String::New("OLEInt32 source type OCVariant is not VT_I4 nor VT_INT nor VT_UI4 nor VT_UINT")));
  DISPFUNCOUT();
  return scope.Close(Int32::New(ocv->v.lVal));
}

Handle<Value> V8Variant::OLEInt64(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  CHECK_OCV(ocv);
  if(ocv->v.vt != VT_I8 && ocv->v.vt != VT_UI8)
    return ThrowException(Exception::TypeError(
      String::New("OLEInt64 source type OCVariant is not VT_I8 nor VT_UI8")));
  DISPFUNCOUT();
#if(0) // may not be supported node.js / v8
  return scope.Close(Int64::New(ocv->v.llVal));
#else
  return scope.Close(Number::New(ocv->v.llVal));
#endif
}

Handle<Value> V8Variant::OLENumber(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  CHECK_OCV(ocv);
  if(ocv->v.vt != VT_R8)
    return ThrowException(Exception::TypeError(
      String::New("OLENumber source type OCVariant is not VT_R8")));
  DISPFUNCOUT();
  return scope.Close(Number::New(ocv->v.dblVal));
}

Handle<Value> V8Variant::OLEDate(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  CHECK_OCV(ocv);
  if(ocv->v.vt != VT_DATE)
    return ThrowException(Exception::TypeError(
      String::New("OLEDate source type OCVariant is not VT_DATE")));
  SYSTEMTIME syst;
  VariantTimeToSystemTime(ocv->v.date, &syst);
  struct tm t = {0}; // set t.tm_isdst = 0
  t.tm_year = syst.wYear - 1900;
  t.tm_mon = syst.wMonth - 1;
  t.tm_mday = syst.wDay;
  t.tm_hour = syst.wHour;
  t.tm_min = syst.wMinute;
  t.tm_sec = syst.wSecond;
  DISPFUNCOUT();
  return scope.Close(Date::New(mktime(&t) * 1000LL + syst.wMilliseconds));
}

Handle<Value> V8Variant::OLEUtf8(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  CHECK_OCV(ocv);
  if(ocv->v.vt != VT_BSTR)
    return ThrowException(Exception::TypeError(
      String::New("OLEUtf8 source type OCVariant is not VT_BSTR")));
  Handle<Value> result;
  if(!ocv->v.bstrVal) result = Undefined(); // or Null();
  else result = String::New(MBCS2UTF8(BSTR2MBCS(ocv->v.bstrVal)).c_str());
  DISPFUNCOUT();
  return scope.Close(result);
}

Handle<Value> V8Variant::OLEValue(const Arguments& args)
{
  HandleScope scope;
  OLETRACEIN();
  OLETRACEVT(args.This());
  OLETRACEFLUSH();
  Local<Object> thisObject = args.This();
  OLE_PROCESS_CARRY_OVER(thisObject);
  OLETRACEVT(thisObject);
  OLETRACEFLUSH();
  OCVariant *ocv = castedInternalField<OCVariant>(thisObject);
  if(!ocv){ std::cerr << "ocv is null"; std::cerr.flush(); }
  CHECK_OCV(ocv);
  Handle<Value> result = Undefined();
  if(ocv->v.vt == VT_EMPTY) ; // do nothing
  else if(ocv->v.vt == VT_NULL) result = Null();
  else if(ocv->v.vt == VT_DISPATCH) result = thisObject; // through it
  else if(ocv->v.vt == VT_BOOL) result = OLEBoolean(args);
  else if(ocv->v.vt == VT_I4 || ocv->v.vt == VT_INT
  || ocv->v.vt == VT_UI4 || ocv->v.vt == VT_UINT) result = OLEInt32(args);
  else if(ocv->v.vt == VT_I8 || ocv->v.vt == VT_UI8) result = OLEInt64(args);
  else if(ocv->v.vt == VT_R8) result = OLENumber(args);
  else if(ocv->v.vt == VT_DATE) result = OLEDate(args);
  else if(ocv->v.vt == VT_BSTR) result = OLEUtf8(args);
  else if(ocv->v.vt == VT_ARRAY || ocv->v.vt == VT_SAFEARRAY){
    std::cerr << "[Array (not implemented now)]" << std::endl;
    std::cerr.flush();
  }else{
    Handle<Value> s = INSTANCE_CALL(thisObject, "vtName", 0, NULL);
    std::cerr << "[unknown type " << ocv->v.vt << ":" << *String::Utf8Value(s);
    std::cerr << " (not implemented now)]" << std::endl;
    std::cerr.flush();
  }
done:
  OLETRACEOUT();
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
#if(0) // needless to do (instance has been already wrapped in New)
  V8Variant *v = new V8Variant(); // must catch exception
  v->Wrap(instance); // InternalField[0]
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
  CHECK_OCV(ocv);
  Local<Object> thisObject = args.This();
  V8Variant *v = new V8Variant(); // must catch exception
  v->Wrap(thisObject); // InternalField[0]
  thisObject->SetInternalField(1, External::New(ocv));
  Persistent<Object> objectDisposer = Persistent<Object>::New(thisObject);
  objectDisposer.MakeWeak(ocv, Dispose);
  DISPFUNCOUT();
  return args.This();
}

Handle<Value> V8Variant::OLEFlushCarryOver(Handle<Value> v)
{
  HandleScope scope;
  OLETRACEIN();
  OLETRACEVT(v->ToObject());
  Handle<Value> result = Undefined();
  V8Variant *v8v = ObjectWrap::Unwrap<V8Variant>(v->ToObject());
  if(v8v->property_carryover.empty()){
    std::cerr << " *** carryover empty *** " << __FUNCTION__ << std::endl;
    std::cerr.flush();
    // *** or throw exception
  }else{
    const char *name = v8v->property_carryover.c_str();
    {
      OLETRACEPREARGV(String::NewSymbol(name));
      OLETRACEARGV();
    }
    OLETRACEFLUSH();
    Handle<Value> argv[] = {String::NewSymbol(name), Array::New(0)};
    int argc = sizeof(argv) / sizeof(argv[0]);
    v8v->property_carryover.erase();
    result = INSTANCE_CALL(v->ToObject(), "call", argc, argv);
    if(!result->IsObject()){
      OCVariant *rv = V8Variant::CreateOCVariant(result);
      CHECK_OCV(rv);
      OCVariant *o = castedInternalField<OCVariant>(v->ToObject());
      CHECK_OCV(o);
      *o = *rv; // copy and don't delete rv
    }
  }
  OLETRACEOUT();
  return scope.Close(result);
}

Handle<Value> V8Variant::OLEInvoke(bool isCall, const Arguments& args)
{
  HandleScope scope;
  OLETRACEIN();
  OLETRACEVT(args.This());
  OLETRACEARGS();
  OLETRACEFLUSH();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  CHECK_OCV(ocv);
  Handle<Value> av0, av1;
  CHECK_OLE_ARGS(args, 1, av0, av1);
  OCVariant *argchain = NULL;
  Array *a = Array::Cast(*av1);
  for(size_t i = 0; i < a->Length(); ++i){
    OCVariant *o = V8Variant::CreateOCVariant(
      ARRAY_AT(a, (i ? i : a->Length()) - 1));
    CHECK_OCV(o);
    if(!i) argchain = o;
    else argchain->push(o);
  }
  Handle<Object> vResult = V8Variant::CreateUndefined();
  String::Utf8Value u8s(av0);
  wchar_t *wcs = u8s2wcs(*u8s);
  if(!wcs && argchain) delete argchain;
  BEVERIFY(done, wcs);
  try{
    OCVariant *rv = isCall ? // argchain will be deleted automatically
      ocv->invoke(wcs, argchain, true) : ocv->getProp(wcs, argchain);
    if(rv){
      OCVariant *o = castedInternalField<OCVariant>(vResult);
      CHECK_OCV(o);
      *o = *rv; // copy and don't delete rv
    }
  }catch(OLE32coreException e){ std::cerr << e.errorMessage(*u8s); goto done;
  }catch(char *e){ std::cerr << e << *u8s << std::endl; goto done;
  }
  free(wcs); // *** it may leak when error ***
#if(0) // It does not work. (when access to NamedPropertyHandler)
  Handle<Value> result = vResult;
#else // It works.
  Handle<Value> result = INSTANCE_CALL(vResult, "toValue", 0, NULL);
#endif
  OLETRACEOUT();
  return scope.Close(result);
done:
  OLETRACEOUT();
  return ThrowException(Exception::TypeError(
    String::New(__FUNCTION__" failed")));
}

Handle<Value> V8Variant::OLECall(const Arguments& args)
{
  HandleScope scope;
  OLETRACEIN();
  OLETRACEVT(args.This());
  OLETRACEARGS();
  OLETRACEFLUSH();
  Handle<Value> r = V8Variant::OLEInvoke(true, args); // as Call
  OLETRACEOUT();
  return scope.Close(r);
}

Handle<Value> V8Variant::OLEGet(const Arguments& args)
{
  HandleScope scope;
  OLETRACEIN();
  OLETRACEVT(args.This());
  OLETRACEARGS();
  OLETRACEFLUSH();
  Handle<Value> r = V8Variant::OLEInvoke(false, args); // as Get
  OLETRACEOUT();
  return scope.Close(r);
}

Handle<Value> V8Variant::OLESet(const Arguments& args)
{
  HandleScope scope;
  OLETRACEIN();
  OLETRACEVT(args.This());
  OLETRACEARGS();
  OLETRACEFLUSH();
  Local<Object> thisObject = args.This();
  OLE_PROCESS_CARRY_OVER(thisObject);
  OCVariant *ocv = castedInternalField<OCVariant>(thisObject);
  CHECK_OCV(ocv);
  Handle<Value> av0, av1;
  CHECK_OLE_ARGS(args, 2, av0, av1);
  OCVariant *argchain = V8Variant::CreateOCVariant(av1);
  if(!argchain)
    return ThrowException(Exception::TypeError(String::New(
      __FUNCTION__" the second argument is not valid (null OCVariant)")));
  bool result = false;
  String::Utf8Value u8s(av0);
  wchar_t *wcs = u8s2wcs(*u8s);
  BEVERIFY(done, wcs);
  try{
    ocv->putProp(wcs, argchain); // argchain will be deleted automatically
  }catch(OLE32coreException e){ std::cerr << e.errorMessage(*u8s); goto done;
  }catch(char *e){ std::cerr << e << *u8s << std::endl; goto done;
  }
  free(wcs); // *** it may leak when error ***
  result = true;
  OLETRACEOUT();
  return scope.Close(Boolean::New(result));
done:
  OLETRACEOUT();
  return ThrowException(Exception::TypeError(
    String::New(__FUNCTION__" failed")));
}

Handle<Value> V8Variant::OLECallComplete(const Arguments& args)
{
  HandleScope scope;
  OLETRACEIN();
  OLETRACEVT(args.This());
  Handle<Value> result = Undefined();
  V8Variant *v8v = ObjectWrap::Unwrap<V8Variant>(args.This());
  if(v8v->property_carryover.empty()){
    std::cerr << " *** carryover empty *** " << __FUNCTION__ << std::endl;
    std::cerr.flush();
    // *** or throw exception
  }else{
    const char *name = v8v->property_carryover.c_str();
    {
      OLETRACEPREARGV(String::NewSymbol(name));
      OLETRACEARGV();
    }
    OLETRACEARGS();
    OLETRACEFLUSH();
    Handle<Array> a = Array::New(args.Length());
    for(int i = 0; i < args.Length(); ++i) ARRAY_SET(a, i, args[i]);
    Handle<Value> argv[] = {String::NewSymbol(name), a};
    int argc = sizeof(argv) / sizeof(argv[0]);
    v8v->property_carryover.erase();
    result = INSTANCE_CALL(args.This(), "call", argc, argv);
  }
//_
//Handle<Value> r = INSTANCE_CALL(Handle<Object>::Cast(v), "toValue", 0, NULL);
  OLETRACEOUT();
  return scope.Close(result);
}

Handle<Value> V8Variant::OLEGetAttr(
  Local<String> name, const AccessorInfo& info)
{
  HandleScope scope;
  OLETRACEIN();
  OLETRACEVT(info.This());
  {
    OLETRACEPREARGV(name);
    OLETRACEARGV();
  }
  OLETRACEFLUSH();
  String::Utf8Value u8name(name);
  Local<Object> thisObject = info.This();

  // Why GetAttr comes twice for () in the third loop instead of CallComplete ?
  // Because of the Crankshaft v8's run-time optimizer ?
  {
    V8Variant *v8v = ObjectWrap::Unwrap<V8Variant>(thisObject);
    if(!v8v->property_carryover.empty()){
      if(v8v->property_carryover == *u8name){
        OLETRACEOUT();
        return scope.Close(thisObject); // through it
      }
    }
  }

#if(0)
  if(std::string("call") == *u8name || std::string("get") == *u8name
  || std::string("_") == *u8name || std::string("toValue") == *u8name
//|| std::string("valueOf") == *u8name || std::string("toString") == *u8name
  ){
    OLE_PROCESS_CARRY_OVER(thisObject);
  }
#else
  if(std::string("set") != *u8name
  && std::string("toBoolean") != *u8name
  && std::string("toInt32") != *u8name && std::string("toInt64") != *u8name
  && std::string("toNumber") != *u8name && std::string("toDate") != *u8name
  && std::string("toUtf8") != *u8name
  && std::string("inspect") != *u8name && std::string("constructor") != *u8name
  && std::string("valueOf") != *u8name && std::string("toString") != *u8name
  && std::string("toLocaleString") != *u8name
  && std::string("hasOwnProperty") != *u8name
  && std::string("isPrototypeOf") != *u8name
  && std::string("propertyIsEnumerable") != *u8name
//&& std::string("_") != *u8name
  ){
    OLE_PROCESS_CARRY_OVER(thisObject);
  }
#endif
  OLETRACEVT(thisObject);
  // Can't use INSTANCE_CALL here. (recursion itself)
  // So it returns Object's fundamental function and custom function:
  //   inspect ?, constructor valueOf toString toLocaleString
  //   hasOwnProperty isPrototypeOf propertyIsEnumerable
  static fundamental_attr fundamentals[] = {
    {0, "call", OLECall}, {0, "get", OLEGet}, {0, "set", OLESet},
    {0, "isA", OLEIsA}, {0, "vtName", OLEVTName}, // {"vt_names", ???},
    {!0, "toBoolean", OLEValue},
    {!0, "toInt32", OLEValue}, {!0, "toInt64", OLEValue},
    {!0, "toNumber", OLEValue}, {!0, "toDate", OLEValue},
    {!0, "toUtf8", OLEValue},
    {0, "toValue", OLEValue},
    {0, "inspect", NULL}, {0, "constructor", NULL}, {0, "valueOf", OLEValue},
    {0, "toString", OLEValue}, {0, "toLocaleString", OLEValue},
    {0, "hasOwnProperty", NULL}, {0, "isPrototypeOf", NULL},
    {0, "propertyIsEnumerable", NULL}
  };
  for(int i = 0; i < sizeof(fundamentals) / sizeof(fundamentals[0]); ++i){
    if(std::string(fundamentals[i].name) != *u8name) continue;
    if(fundamentals[i].obsoleted){
      std::cerr << " *** ## [." << fundamentals[i].name;
      std::cerr << "()] is obsoleted. ## ***" << std::endl;
      std::cerr.flush();
    }
    OLETRACEFLUSH();
    OLETRACEOUT();
    return scope.Close(FunctionTemplate::New(
      fundamentals[i].func, thisObject)->GetFunction());
  }
  if(std::string("_") == *u8name){ // through it when "_"
#if(0)
    std::cerr << " *** ## [._] is obsoleted. ## ***" << std::endl;
    std::cerr.flush();
#endif
  }else{
    Handle<Object> vResult = V8Variant::CreateUndefined(); // uses much memory
    OCVariant *rv = V8Variant::CreateOCVariant(thisObject);
    CHECK_OCV(rv);
    OCVariant *o = castedInternalField<OCVariant>(vResult);
    CHECK_OCV(o);
    *o = *rv; // copy and don't delete rv
    V8Variant *v8v = ObjectWrap::Unwrap<V8Variant>(vResult);
    v8v->property_carryover.assign(*u8name);
    OLETRACEPREARGV(name);
    OLETRACEARGV();
    OLETRACEFLUSH();
    OLETRACEOUT();
    return scope.Close(vResult); // convert and hold it (uses much memory)
  }
  OLETRACEFLUSH();
  OLETRACEOUT();
  return scope.Close(thisObject); // through it
}

Handle<Value> V8Variant::OLESetAttr(
  Local<String> name, Local<Value> val, const AccessorInfo& info)
{
  HandleScope scope;
  OLETRACEIN();
  OLETRACEVT(info.This());
  Handle<Value> argv[] = {name, val};
  int argc = sizeof(argv) / sizeof(argv[0]);
  OLETRACEARGV();
  OLETRACEFLUSH();
  Handle<Value> r = INSTANCE_CALL(info.This(), "set", argc, argv);
  OLETRACEOUT();
  return scope.Close(r);
}

Handle<Value> V8Variant::Finalize(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
#if(0)
  std::cerr << __FUNCTION__ << " Finalizer is called\a" << std::endl;
  std::cerr.flush();
#endif
  Local<Object> thisObject = args.This();
#if(0)
  V8Variant *v = ObjectWrap::Unwrap<V8Variant>(thisObject);
  if(v) delete v; // it has been already deleted ?
  thisObject->SetInternalField(0, External::New(NULL));
#endif
#if(1) // now GC will call Disposer automatically
  OCVariant *ocv = castedInternalField<OCVariant>(thisObject);
  if(ocv) delete ocv;
#endif
  thisObject->SetInternalField(1, External::New(NULL));
  DISPFUNCOUT();
  return args.This();
}

void V8Variant::Dispose(Persistent<Value> handle, void *param)
{
  DISPFUNCIN();
#if(0)
//  std::cerr << __FUNCTION__ << " Disposer is called\a" << std::endl;
  std::cerr << __FUNCTION__ << " Disposer is called" << std::endl;
  std::cerr.flush();
#endif
  Local<Object> thisObject = handle->ToObject();
#if(0) // it has been already deleted ?
  V8Variant *v = ObjectWrap::Unwrap<V8Variant>(thisObject);
  if(!v){
    std::cerr << __FUNCTION__;
    std::cerr << "InternalField[0] has been already deleted" << std::endl;
    std::cerr.flush();
  }else delete v; // it has been already deleted ?
  BEVERIFY(done, thisObject->InternalFieldCount() > 0);
  thisObject->SetInternalField(0, External::New(NULL));
#endif
  OCVariant *p = castedInternalField<OCVariant>(thisObject);
  if(!p){
    std::cerr << __FUNCTION__;
    std::cerr << "InternalField[1] has been already deleted" << std::endl;
    std::cerr.flush();
  }
//  else{
    OCVariant *ocv = static_cast<OCVariant *>(param); // ocv may be same as p
    if(ocv) delete ocv;
//  }
  BEVERIFY(done, thisObject->InternalFieldCount() > 1);
  thisObject->SetInternalField(1, External::New(NULL));
done:
  handle.Dispose();
  DISPFUNCOUT();
}

void V8Variant::Finalize()
{
  assert(!finalized);
  finalized = true;
}

} // namespace node_win32ole
