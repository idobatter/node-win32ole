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
  // create context
  const char *extensionNames[] = {"v8/gc",};
  ExtensionConfiguration extensions(
    sizeof(extensionNames) / sizeof(extensionNames[0]), extensionNames);
  Handle<ObjectTemplate> global = ObjectTemplate::New();
  // Persistent<Context> context = Context::New(NULL, global);
  Handle<Context> context = Context::New(&extensions, global);
  Context::Scope context_scope(context);
  BDISPFUNCDAT("context %s "__FUNCTION__" %s\n", "preset", "end");
  HandleScope scope;
  BDISPFUNCIN();
  Local<String> sourceObj = String::New("gc()");
  Local<Script> scriptObj = Script::Compile(sourceObj);
  if(scriptObj.IsEmpty())
    return ThrowException(Exception::TypeError(
      String::New("Can't execute v8/gc : gc();")));
  Local<Value> local_result = scriptObj->Run();
  if(local_result.IsEmpty())
    return ThrowException(Exception::TypeError(
      String::New("Can't get result of v8/gc : gc();")));
  String::Utf8Value result(local_result);
  // context.Dispose(); // when Persistent<Context>
  BDISPFUNCOUT();
  return scope.Close(String::New(*result));
}

} // namespace node_win32ole
