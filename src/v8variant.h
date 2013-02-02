#ifndef __V8VARIANT_H__
#define __V8VARIANT_H__

#include "node_win32ole.h"
#include "ole32core.h"

using namespace v8;
using namespace ole32core;

namespace node_win32ole {

class V8Variant : public node::ObjectWrap {
public:
  static Persistent<FunctionTemplate> constructor_template;
  static void Init(Handle<Object> target);
  static Handle<Value> New(const Arguments& args);
  static Handle<Value> Del(const Arguments& args);
public:
  V8Variant() : node::ObjectWrap(), finalized(false) {}
  ~V8Variant() { if(!finalized) Finalize(); }
protected:
  void Finalize();
protected:
  bool finalized;
  OCVariant *ocv;
};

} // namespace node_win32ole

#endif // __V8VARIANT_H__
