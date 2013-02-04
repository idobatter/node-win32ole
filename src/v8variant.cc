/*
  v8variant.cc
*/

#include "v8variant.h"
#include "ole32core.h"

using namespace v8;
using namespace ole32core;

namespace node_win32ole {

Persistent<FunctionTemplate> V8Variant::clazz;

void V8Variant::Init(Handle<Object> target)
{
  HandleScope scope;
  Local<FunctionTemplate> t = FunctionTemplate::New(New);
  clazz = Persistent<FunctionTemplate>::New(t);
  clazz->InstanceTemplate()->SetInternalFieldCount(1);
  clazz->SetClassName(String::NewSymbol("V8Variant"));
  NODE_SET_PROTOTYPE_METHOD(clazz, "New", New);
  NODE_SET_PROTOTYPE_METHOD(clazz, "get", OLEGet);
  NODE_SET_PROTOTYPE_METHOD(clazz, "set", OLESet);
  NODE_SET_PROTOTYPE_METHOD(clazz, "call", OLECall);
  NODE_SET_PROTOTYPE_METHOD(clazz, "Finalize", Finalize);
  target->Set(String::NewSymbol("V8Variant"), clazz->GetFunction());
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
  if(!ocv)
    return ThrowException(Exception::TypeError(
      String::New("Can't create new V8Variant object (null OCVariant)")));
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
  if(!ocv)
    return ThrowException(Exception::TypeError(
      String::New("Can't access to V8Variant object (null OCVariant)")));
  DISPFUNCOUT();
  return args.This();
}

Handle<Value> V8Variant::OLESet(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  if(!ocv)
    return ThrowException(Exception::TypeError(
      String::New("Can't access to V8Variant object (null OCVariant)")));
  if(args.Length() < 2)
    return ThrowException(Exception::TypeError(
      String::New("it takes 2 arguments")));
  if(!args[0]->IsString())
    return ThrowException(Exception::TypeError(
      String::New("first argument is not a Symbol")));

#if(0)
  if(!args[1]->IsInt32())
    return ThrowException(Exception::TypeError(
      String::New("second argument is not a Int32")));
#endif
  if(!args[1]->IsBoolean())
    return ThrowException(Exception::TypeError(
      String::New("second argument is not a Boolean")));

  bool result = false;
  String::Utf8Value u8s(args[0]);
  wchar_t *wcs = u8s2wcs(*u8s);
  BEVERIFY(done, wcs);
  try{
    ocv->putProp(wcs, new OCVariant((long)(args[1]->BooleanValue() ? 1 : 0)));
  }catch(OLE32coreException e){ std::cerr << e.errorMessage(*u8s); goto done;
  }catch(char *e){ std::cerr << e << *u8s << std::endl; goto done;
  }
  free(wcs);
  result = true;
done:
  DISPFUNCOUT();
  return scope.Close(Boolean::New(result));
}

Handle<Value> V8Variant::OLECall(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
  OCVariant *ocv = castedInternalField<OCVariant>(args.This());
  if(!ocv)
    return ThrowException(Exception::TypeError(
      String::New("Can't access to V8Variant object (null OCVariant)")));
  DISPFUNCOUT();
  return args.This();
}

Handle<Value> V8Variant::Finalize(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
  Local<Object> thisObject = args.This();
#if(0) // now GC will call Disposer automatically
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
  OCVariant *ocv = static_cast<OCVariant *>(param);
  if(ocv) delete ocv;
  // thisObject->SetInternalField(0, External::New(NULL));
  handle.Dispose();
  DISPFUNCOUT();
}

void V8Variant::Finalize()
{
  assert(!finalized);
  finalized = true;
}

} // namespace node_win32ole
