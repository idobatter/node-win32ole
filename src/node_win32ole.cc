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
  char *u8s = *s;
  size_t u8len = strlen(u8s);
  size_t wclen = MultiByteToWideChar(CP_UTF8, 0, u8s, u8len, NULL, 0);
  wchar_t *wcs = (wchar_t *)malloc((wclen + 1) * sizeof(wchar_t));
  wclen = MultiByteToWideChar(CP_UTF8, 0, u8s, u8len, wcs, wclen + 1);
  wcs[wclen] = L'\0';
#ifdef DEBUG
  fwprintf(stderr, L"wcs: %s\n", wcs);
#endif
  CLSID clsid;
  HRESULT hr = CLSIDFromProgID(wcs, &clsid);
  free(wcs);
  if(FAILED(hr)){
    fprintf(stderr, "error: CLSIDFromProgID() %s\n", u8s);
    return scope.Close(Boolean::New(result));
  }
  hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch,
    (void **)&vDisp.pdispVal);
  if(FAILED(hr)){
    fprintf(stderr, "GetLastError: %08x\n", GetLastError()); // 000003f0
    fprintf(stderr, "error: CoCreateInstance %s\n", u8s); // no token
    fprintf(stderr, "clsid:"); // 00024500-0000-0000-c000-000000000046 (Excel)
    for(int i = 0; i < sizeof(CLSID); ++i)
      fprintf(stderr, " %02x", ((unsigned char *)&clsid)[i]);
    fprintf(stderr, "\n");
    return scope.Close(Boolean::New(result));
  }
  result = true;
  return scope.Close(Boolean::New(result));
}

Handle<Value> Statement::Finalize(const Arguments& args)
{
  HandleScope scope;
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
