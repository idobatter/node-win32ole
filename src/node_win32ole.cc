/*
  node_win32ole.cc
  $ cp ~/.node-gyp/0.8.18/ia32/node.lib to ~/.node-gyp/0.8.18/node.lib
  $ node-gyp configure
  $ node-gyp build
  $ node test/init_win32ole.test.js
*/

#include "node_win32ole.h"
#include "client.h"
#include "v8variant.h"

using namespace v8;

namespace node_win32ole {

Persistent<Object> module_target;

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

Handle<Value> Method_printACP(const Arguments& args) // UTF-8 to MBCS (.ACP)
{
  HandleScope scope;
  if(args.Length() >= 1){
    String::Utf8Value s(args[0]);
    char *cs = *s;
    printf(UTF82MBCS(std::string(cs)).c_str());
  }
  return scope.Close(Boolean::New(true));
}

Handle<Value> Method_print(const Arguments& args) // through (as ASCII)
{
  HandleScope scope;
  if(args.Length() >= 1){
    String::Utf8Value s(args[0]);
    char *cs = *s;
    printf(cs); // printf("%s\n", cs);
  }
  return scope.Close(Boolean::New(true));
}

} // namespace node_win32ole

using namespace node_win32ole;

namespace {

void init(Handle<Object> target)
{
  module_target = Persistent<Object>::New(target);
  V8Variant::Init(target);
  Client::Init(target);
  target->Set(String::NewSymbol("VERSION"),
    String::New("0.0.0 (will be set later)"),
    static_cast<PropertyAttribute>(DontDelete));
  target->Set(String::NewSymbol("MODULEDIRNAME"),
    String::New("/tmp"),
    static_cast<PropertyAttribute>(DontDelete));
  target->Set(String::NewSymbol("SOURCE_TIMESTAMP"),
    String::NewSymbol(__FILE__ " " __DATE__ " " __TIME__),
    static_cast<PropertyAttribute>(ReadOnly | DontDelete));
  target->Set(String::NewSymbol("version"),
    FunctionTemplate::New(Method_version)->GetFunction());
  target->Set(String::NewSymbol("printACP"),
    FunctionTemplate::New(Method_printACP)->GetFunction());
  target->Set(String::NewSymbol("print"),
    FunctionTemplate::New(Method_print)->GetFunction());
  target->Set(String::NewSymbol("gettimeofday"),
    FunctionTemplate::New(Method_gettimeofday)->GetFunction());
  target->Set(String::NewSymbol("sleep"),
    FunctionTemplate::New(Method_sleep)->GetFunction());
  target->Set(String::NewSymbol("force_gc_extension"),
    FunctionTemplate::New(Method_force_gc_extension)->GetFunction());
  target->Set(String::NewSymbol("force_gc_internal"),
    FunctionTemplate::New(Method_force_gc_internal)->GetFunction());
}

} // namespace

NODE_MODULE(node_win32ole, init)
