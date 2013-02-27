#ifndef __NODE_WIN32OLE_H__
#define __NODE_WIN32OLE_H__

#include <node_buffer.h>
#include <node.h>
#include <v8.h>

using namespace v8;

namespace node_win32ole {

#define CHECK_OCV(ocv) do{ \
    if(!(ocv)) \
      return ThrowException(Exception::TypeError(String::New( \
        __FUNCTION__" can't access to V8Variant (null OCVariant)"))); \
  }while(0)

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
    Handle<Value> r = V8Variant::OLEFlushCarryOver(th); \
    if(!r->IsObject()){ \
      std::cerr << "** CarryOver primitive ** " << __FUNCTION__ << std::endl; \
      std::cerr.flush(); \
      return scope.Close(r); \
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
#define OLE_PROCESS_CARRY_OVER(th) do{ \
    V8Variant *v8v = ObjectWrap::Unwrap<V8Variant>(th); \
    if(v8v->property_carryover.empty()) break; \
    Handle<Value> r = V8Variant::OLEFlushCarryOver(th); \
    if(!r->IsObject()) return scope.Close(r); \
    th = r->ToObject(); \
  }while(0)
#endif

#define GET_PROP(obj, prop) (obj)->Get(String::NewSymbol(prop))

#define ARRAY_AT(a, i) (a)->Get(String::NewSymbol(to_s(i).c_str()))
#define ARRAY_SET(a, i, v) (a)->Set(String::NewSymbol(to_s(i).c_str()), (v))

#define INSTANCE_CALL(obj, method, argc, argv) Handle<Function>::Cast( \
  GET_PROP((obj), (method)))->Call((obj), (argc), (argv))

template <class T> T *castedInternalField(Handle<Object> object, int fidx=1)
{
  return static_cast<T *>(
    Local<External>::Cast(object->GetInternalField(fidx))->Value());
}

extern Persistent<Object> module_target;

Handle<Value> Method_version(const Arguments& args);
Handle<Value> Method_printACP(const Arguments& args); // UTF-8 to MBCS (.ACP)
Handle<Value> Method_print(const Arguments& args); // through (as ASCII)
Handle<Value> Method_gettimeofday(const Arguments& args);
Handle<Value> Method_sleep(const Arguments& args); // ms, bool: msg, bool: \n
Handle<Value> Method_force_gc_extension(const Arguments& args); // v8/gc : gc()
Handle<Value> Method_force_gc_internal(const Arguments& args); // v8/src/v8.h

} // namespace node_win32ole

#endif // __NODE_WIN32OLE_H__
