#ifndef __STATEMENT_H__
#define __STATEMENT_H__

#include "node_win32ole.h"
#include "ole32core.h"

using namespace v8;
using namespace ole32core;

namespace node_win32ole {

class Statement : public node::ObjectWrap {
public:
  static Persistent<FunctionTemplate> clazz;
  static void Init(Handle<Object> target);
  static Handle<Value> New(const Arguments& args);
  static Handle<Value> Dispatch(const Arguments& args);
  static Handle<Value> Finalize(const Arguments& args);
public:
  Statement() : node::ObjectWrap(), finalized(false) {}
  ~Statement() { if(!finalized) Finalize(); }
protected:
  void Finalize();
protected:
  bool finalized;
  static OLE32core oc;
};

} // namespace node_win32ole

#endif // __STATEMENT_H__
