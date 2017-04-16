/*
  force_gc_internal.cc
*/

#include "node_win32ole.h"
#include "ole32core.h"
#include <iostream>

using namespace v8;

namespace node_win32ole {

    void Method_force_gc_internal(const FunctionCallbackInfo<Value>& args) // v8/src/v8.h
    {
        Isolate *isolate = args.GetIsolate();
        std::cerr << "-in: " __FUNCTION__ << std::endl;
        if (args.Length() < 1) RETURN_TYPE_ERROR("this function takes at least 1 argument(s)")
        if (!args[0]->IsInt32()) RETURN_TYPE_ERROR("type of argument 1 must be Int32")
        int flags = (int)args[0]->Int32Value();

        //while (!v8::V8::IdleNotification()) {}
        //isolate->IdleNotificationDeadline();   
        RETURN_ERROR("function not implemented, becose IdleNotification is depricated");

        std::cerr << "-out: " __FUNCTION__ << std::endl;
        args.GetReturnValue().Set(Boolean::New(isolate, true));
    }

} // namespace node_win32ole
