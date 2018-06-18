#ifndef __NODE_WIN32OLE_H__
#define __NODE_WIN32OLE_H__

#include <v8.h>
#include <node.h>
#include <node_buffer.h>
#include <node_object_wrap.h>

using namespace v8;

namespace node_win32ole {

#define NODE_DEFINE_PROP(target, key, val) NODE_DEFINE_PROP_(target, key, val, 0) 
#define NODE_DEFINE_PROP_READONLY(target, key, val) NODE_DEFINE_PROP_(target, key, val, ReadOnly) 
#define NODE_DEFINE_PROP_(target, key, val, opt) { \
    Isolate* isolate = target->GetIsolate(); \
    Local<Context> context = isolate->GetCurrentContext(); \
    target->DefineOwnProperty(context, String::NewFromUtf8(isolate, key), \
        String::NewFromUtf8(isolate, val), \
        static_cast<PropertyAttribute>(opt | DontDelete)); \
}

#define THROW_ERROR_T(terr, msg) isolate->ThrowException(terr(String::NewFromUtf8(isolate, msg)))
#define THROW_ERROR(msg) RETURN_ERROR_T(Exception::Error, msg)
#define THROW_TYPE_ERROR(msg) RETURN_ERROR_T(Exception::TypeError, msg)
#define THROW_RANGE_ERROR(msg) RETURN_ERROR_T(Exception::TangeError, msg)

#define RETURN_ERROR_T(terr, msg) { THROW_ERROR_T(terr, msg); return; }
#define RETURN_ERROR(msg) RETURN_ERROR_T(Exception::Error, msg)
#define RETURN_TYPE_ERROR(msg) RETURN_ERROR_T(Exception::TypeError, msg)
#define RETURN_RANGE_ERROR(msg) RETURN_ERROR_T(Exception::RangeError, msg)

    
#define CHECK_OCV(ocv) do { \
    if (!(ocv)) RETURN_ERROR(__FUNCTION__" can't access to V8Variant (null OCVariant)") \
  } while (0)

#if(0)
#define OLETRACEIN() BDISPFUNCIN()
#define OLETRACEVT(th) do{ \
    OCVariant *ocv = castedInternalField<OCVariant>(th); \
    if(!ocv){ std::cerr << "*** OCVariant is NULL ***"; std::cerr.flush(); } \
    CHECK_OCV(ocv); \
    std::cerr << "0x" << std::setw(8) << std::left << std::hex << ocv << ":"; \
    std::cerr << "vt=" << ocv->v.vt << ":"; \
    std::cerr.flush(); \
  }while(0)
#define OLETRACEARG(v) do{ \
    std::cerr << (v->IsObject() ? "OBJECT" : *String::Utf8Value(v)) << ","; \
  }while(0)
#define OLETRACEPREARGV(sargs) Handle<Value> argv[] = { sargs }; \
  int argc = sizeof(argv) / sizeof(argv[0])
#define OLETRACEARGV() do{ \
    for(int i = 0; i < argc; ++i) OLETRACEARG(argv[i]); \
  }while(0)
#define OLETRACEARGS() do{ \
    for(int i = 0; i < args.Length(); ++i) OLETRACEARG(args[i]); \
  }while(0)
#define OLETRACEFLUSH() do{ std::cerr<<std::endl; std::cerr.flush(); }while(0)
#define OLETRACEOUT() BDISPFUNCOUT()
#define OLE_PROCESS_CARRY_OVER(th) do{ \
    V8Variant *v8v = ObjectWrap::Unwrap<V8Variant>(th); \
    if(v8v->property_carryover.empty()) break; \
    Handle<Value> r; V8Variant::OLEFlushCarryOver(isolate, th, r); \
    if(!r->IsObject()){ \
      std::cerr << "** CarryOver primitive ** " << __FUNCTION__ << std::endl; \
      std::cerr.flush(); \
      args.GetReturnValue().Set(r); \
      return; \
    } \
    th = r->ToObject(); \
  }while(0)
#else
#define OLETRACEIN()
#define OLETRACEVT(th)
#define OLETRACEARG(v)
#define OLETRACEPREARGV(sargs)
#define OLETRACEARGV()
#define OLETRACEARGS()
#define OLETRACEFLUSH()
#define OLETRACEOUT()
#define OLE_PROCESS_CARRY_OVER(th) do { \
    V8Variant *v8v = ObjectWrap::Unwrap<V8Variant>(th); \
    if (v8v->property_carryover.empty()) break; \
    Handle<Value> r; V8Variant::OLEFlushCarryOver(isolate, th, r); \
    if (!r->IsObject()) { args.GetReturnValue().Set(r); return; }\
    th = r->ToObject(); \
  } while(0)
#endif

#define GET_PROP(obj, prop) (obj)->Get(String::NewFromUtf8(isolate, prop))

#define ARRAY_AT(a, i) (a)->Get(String::NewFromUtf8(isolate, to_s(i).c_str()))
#define ARRAY_SET(a, i, v) (a)->Set(String::NewFromUtf8(isolate, to_s(i).c_str()), (v))

#define INSTANCE_CALL(obj, method, argc, argv) Handle<Function>::Cast( \
  GET_PROP((obj), (method)))->Call((obj), (argc), (argv))

template <class T> T *castedInternalField(Handle<Object> object, int fidx=1)
{
  return static_cast<T *>(
    Local<External>::Cast(object->GetInternalField(fidx))->Value());
}

extern Persistent<Object> module_target;

void Method_version(const FunctionCallbackInfo<Value>& args);
void Method_printACP(const FunctionCallbackInfo<Value>& args); // UTF-8 to MBCS (.ACP)
void Method_print(const FunctionCallbackInfo<Value>& args); // through (as ASCII)
void Method_gettimeofday(const FunctionCallbackInfo<Value>& args);
void Method_sleep(const FunctionCallbackInfo<Value>& args); // ms, bool: msg, bool: \n
void Method_force_gc_extension(const FunctionCallbackInfo<Value>& args); // v8/gc : gc()
void Method_force_gc_internal(const FunctionCallbackInfo<Value>& args); // v8/src/v8.h

} // namespace node_win32ole

#endif // __NODE_WIN32OLE_H__
