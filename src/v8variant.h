#ifndef __V8VARIANT_H__
#define __V8VARIANT_H__

#include "node_win32ole.h"
#include "ole32core.h"

using namespace v8;
using namespace ole32core;

namespace node_win32ole {

class V8Variant : public node::ObjectWrap {
public:
  static Persistent<FunctionTemplate> clazz;
  static void Init(Handle<Object> target);
  static std::string CreateStdStringMBCSfromUTF8(Handle<Value> v);
  static OCVariant *CreateOCVariant(Handle<Value> v);
  static Handle<Value> OLEBoolean(const Arguments& args);
  static Handle<Value> OLEInt32(const Arguments& args);
  static Handle<Value> OLENumber(const Arguments& args);
  static Handle<Value> OLEUtf8(const Arguments& args);
  static Handle<Object> CreateUndefined(void);
  static Handle<Value> New(const Arguments& args);
  static Handle<Value> OLEGet(const Arguments& args);
  static Handle<Value> OLESet(const Arguments& args);
  static Handle<Value> OLECall(const Arguments& args);
  static Handle<Value> Finalize(const Arguments& args);
public:
  V8Variant() : node::ObjectWrap(), finalized(false) {}
  ~V8Variant() { BDISPFUNCIN(); if(!finalized) Finalize(); BDISPFUNCOUT(); }
protected:
  static void Dispose(Persistent<Value> handle, void *param);
  void Finalize();
protected:
  bool finalized;
};

} // namespace node_win32ole

#endif // __V8VARIANT_H__
