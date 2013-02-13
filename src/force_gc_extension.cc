/*
  force_gc_extension.cc
*/

#include "node_win32ole.h"
#include "ole32core.h"

using namespace v8;
using namespace ole32core;

namespace node_win32ole {

Handle<Value> Method_force_gc_extension(const Arguments& args) // v8/gc : gc()
{
  BDISPFUNCDAT("context %s "__FUNCTION__" %s\n", "preset", "start");
  // create context with extension(s)
  const char *extensionNames[] = {"v8/gc",};
  ExtensionConfiguration extensions(
    sizeof(extensionNames) / sizeof(extensionNames[0]), extensionNames);
  Handle<ObjectTemplate> global = ObjectTemplate::New();
  // another way get 'global' by ( global = context->Global() ) but no context
  // Persistent<Context> context = Context::New(NULL, global);
  Persistent<Context> context = Context::New(&extensions, global);
  Context::Scope context_scope(context);
  BDISPFUNCDAT("context %s "__FUNCTION__" %s\n", "preset", "end");
  HandleScope scope;
  BDISPFUNCIN();
  Local<String> sourceObj = String::New("gc()");
  TryCatch try_catch;
  Local<Script> scriptObj = Script::Compile(sourceObj);
  if(scriptObj.IsEmpty()){
    std::string msg("Can't compile v8/gc : gc();\n");
    String::Utf8Value u8s(try_catch.Exception());
    msg += *u8s;
    return ThrowException(Exception::TypeError(String::New(msg.c_str())));
  }
  Local<Value> local_result = scriptObj->Run();
  if(local_result.IsEmpty()){
    std::string msg("Can't get execution result of v8/gc : gc();\n");
    String::Utf8Value u8s(try_catch.Exception());
    msg += *u8s;
    return ThrowException(Exception::TypeError(String::New(msg.c_str())));
  }
  String::Utf8Value result(local_result);
  context.Dispose(); // dispose Persistent<Context>
  BDISPFUNCOUT();
  return scope.Close(String::New(*result));
}

} // namespace node_win32ole
