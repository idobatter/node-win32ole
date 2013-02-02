#ifndef __NODE_WIN32OLE_H__
#define __NODE_WIN32OLE_H__

#include <node_buffer.h>
#include <node.h>
#include <v8.h>

#include "ole32core.h"

using namespace v8;
using namespace ole32core;

namespace node_win32ole {

Handle<Value> Method_version(const Arguments& args);
Handle<Value> Method_print(const Arguments& args);
Handle<Value> Method_gettimeofday(const Arguments& args);

class Statement : public node::ObjectWrap {
public:
  static Persistent<FunctionTemplate> constructor_template;
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

#endif // __NODE_WIN32OLE_H__
