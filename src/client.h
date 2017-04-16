#ifndef __CLIENT_H__
#define __CLIENT_H__

#include "node_win32ole.h"

using namespace v8;

namespace node_win32ole {

class Client : public node::ObjectWrap {
public:
  static Persistent<FunctionTemplate> clazz;
  static void Init(Handle<Object> target);
  static void New(const FunctionCallbackInfo<Value>& args);
  static void Dispatch(const FunctionCallbackInfo<Value>& args);
  static void Finalize(const FunctionCallbackInfo<Value>& args);
public:
  Client() : node::ObjectWrap(), finalized(false) {}
  ~Client() { if(!finalized) Finalize(); }
protected:
  static void Dispose(Isolate* isolate, Persistent<Value> handle, void *param);
  void Finalize();
protected:
  bool finalized;
};

} // namespace node_win32ole

#endif // __CLIENT_H__
