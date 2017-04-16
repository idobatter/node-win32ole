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

    void Method_version(const FunctionCallbackInfo<Value>& args)
    {
        Isolate* isolate = args.GetIsolate();
        if (args.Length() >= 0) {
            //
        }
#ifdef DEBUG
        node::node_module *nms = 0;// node::get_builtin_module("Array");
        const char *info = nms ? nms->nm_modname : "not found";
#else
        const char *info ="VERSION";
#endif
        args.GetReturnValue().Set(String::NewFromUtf8(isolate, info));
    }

    void Method_printACP(const FunctionCallbackInfo<Value>& args) // UTF-8 to MBCS (.ACP)
    {
        Isolate* isolate = args.GetIsolate();
        if (args.Length() >= 1) {
            String::Utf8Value s(args[0]);
            char *cs = *s;
            printf(UTF82MBCS(std::string(cs)).c_str());
        }
        args.GetReturnValue().Set(Boolean::New(isolate, true));
    }

    void Method_print(const FunctionCallbackInfo<Value>& args) // through (as ASCII)
    {
        Isolate* isolate = args.GetIsolate();
        if (args.Length() >= 1) {
            String::Utf8Value s(args[0]);
            char *cs = *s;
            printf(cs); // printf("%s\n", cs);
        }
        args.GetReturnValue().Set(Boolean::New(isolate, true));
    }

} // namespace node_win32ole

using namespace node_win32ole;

namespace {

    void init(Handle<Object> target)
    {
        Isolate* isolate = target->GetIsolate();

        module_target.Reset(isolate, target);
        V8Variant::Init(target);
        Client::Init(target);

        NODE_DEFINE_PROP(target, "VERSION", "0.0.0");       // replaced from package.json
        NODE_DEFINE_PROP(target, "MODULEDIRNAME", "/tmp");  // replaced for actual path
        NODE_DEFINE_PROP_READONLY(target, "SOURCE_TIMESTAMP", __DATE__ " " __TIME__);

        NODE_SET_METHOD(target, "version", Method_version);
        NODE_SET_METHOD(target, "printACP", Method_printACP);
        NODE_SET_METHOD(target, "print", Method_print);
        NODE_SET_METHOD(target, "gettimeofday", Method_gettimeofday);
        NODE_SET_METHOD(target, "sleep", Method_sleep);
        NODE_SET_METHOD(target, "force_gc_extension", Method_force_gc_extension);
        NODE_SET_METHOD(target, "force_gc_internal", Method_force_gc_internal);
    }

} // namespace

NODE_MODULE(node_win32ole, init)
