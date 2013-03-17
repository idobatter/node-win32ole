#ifndef __V8VARIANT_H__
#define __V8VARIANT_H__

#include "node_win32ole.h"
#include "ole32core.h"

using namespace v8;
using namespace ole32core;

namespace node_win32ole {

typedef struct _fundamental_attr {
  bool obsoleted;
  const char *name;
  Handle<Value> (*func)(const Arguments& args);
} fundamental_attr;

class V8Variant : public node::ObjectWrap {
public:
  static Persistent<FunctionTemplate> clazz;
  static void Init(Handle<Object> target);
  static std::string CreateStdStringMBCSfromUTF8(Handle<Value> v); // *** p.
  static OCVariant *CreateOCVariant(Handle<Value> v); // *** private
  static Handle<Value> OLEIsA(const Arguments& args);
  static Handle<Value> OLEVTName(const Arguments& args);
  static Handle<Value> OLEBoolean(const Arguments& args); // *** p.
  static Handle<Value> OLEInt32(const Arguments& args); // *** p.
  static Handle<Value> OLEInt64(const Arguments& args); // *** p.
  static Handle<Value> OLENumber(const Arguments& args); // *** p.
  static Handle<Value> OLEDate(const Arguments& args); // *** p.
  static Handle<Value> OLEUtf8(const Arguments& args); // *** p.
  static Handle<Value> OLEValue(const Arguments& args);
  static Handle<Object> CreateUndefined(void); // *** private
  static Handle<Value> New(const Arguments& args);
  static Handle<Value> OLEFlushCarryOver(Handle<Value> v); // *** p.
  static Handle<Value> OLEInvoke(bool isCall, const Arguments& args); // *** p.
  static Handle<Value> OLECall(const Arguments& args);
  static Handle<Value> OLEGet(const Arguments& args);
  static Handle<Value> OLESet(const Arguments& args);
  static Handle<Value> OLECallComplete(const Arguments& args); // *** p.
  static Handle<Value> OLEGetAttr(
    Local<String> name, const AccessorInfo& info); // *** p.
  static Handle<Value> OLESetAttr(
    Local<String> name, Local<Value> val, const AccessorInfo& info); // *** p.
  static Handle<Value> Finalize(const Arguments& args);
public:
  V8Variant() : node::ObjectWrap(), finalized(false), property_carryover() {}
  ~V8Variant() { if(!finalized) Finalize(); }
protected:
  static void Dispose(Persistent<Value> handle, void *param);
  void Finalize();
protected:
  bool finalized;
  std::string property_carryover;
};

} // namespace node_win32ole

#endif // __V8VARIANT_H__
