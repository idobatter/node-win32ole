/*
  force_gc_extension.cc
*/

#include "node_win32ole.h"
#include "ole32core.h"

using namespace v8;
using namespace ole32core;

namespace node_win32ole {

    void Method_force_gc_extension(const FunctionCallbackInfo<Value>& args) // v8/gc : gc()
    {
        Isolate *isolate = args.GetIsolate();
        BDISPFUNCDAT("context %s " __FUNCTION__ " %s\n", "preset", "start");
        // create context with extension(s)
        const char *extensionNames[] = { "v8/gc", };
        ExtensionConfiguration extensions(sizeof(extensionNames) / sizeof(extensionNames[0]), extensionNames);

        //Local<Context> context = isolate->GetCurrentContext();
        //Handle<ObjectTemplate> global = context->Global();

        BDISPFUNCDAT("context %s " __FUNCTION__ " %s\n", "preset", "end");
        BDISPFUNCIN();
        Local<String> sourceObj = String::NewFromUtf8(isolate, "gc()");
        TryCatch try_catch;
        Local<Script> scriptObj = Script::Compile(sourceObj);
        if (scriptObj.IsEmpty()) {
            std::string msg("Can't compile v8/gc : gc();\n");
            String::Utf8Value u8s(try_catch.Exception());
            msg += *u8s;
            RETURN_TYPE_ERROR(msg.c_str())
        }
        Local<Value> local_result = scriptObj->Run();
        if (local_result.IsEmpty()) {
            std::string msg("Can't get execution result of v8/gc : gc();\n");
            String::Utf8Value u8s(try_catch.Exception());
            msg += *u8s;
            RETURN_TYPE_ERROR(msg.c_str())
        }

        //String::Utf8Value result(local_result);
        args.GetReturnValue().Set(local_result->ToString());
        BDISPFUNCOUT();
    }

} // namespace node_win32ole
