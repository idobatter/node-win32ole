#ifndef __NODE_WIN32OLE_H__
#define __NODE_WIN32OLE_H__

#include <node_buffer.h>
#include <node.h>
#include <v8.h>

using namespace v8;

Handle<Value> Method_version(const Arguments& args);
Handle<Value> Method_print(const Arguments& args);
Handle<Value> Method_gettimeofday(const Arguments& args);

class Statement : public node::ObjectWrap {
public:
  static Persistent<FunctionTemplate> constructor_template;
  static void Init(Handle<Object> target);
  static Handle<Value> New(const Arguments& args);
  static Handle<Value> Dispatch(const Arguments& args);
public:
  Statement() : node::ObjectWrap(), finalized(false) {
    CoInitialize(NULL);
    vDisp.vt = VT_DISPATCH;
    vDisp.pdispVal = 0;
  }
  ~Statement() {
    if(!finalized){
      Finalize();
      if(vDisp.vt == VT_DISPATCH && vDisp.pdispVal)
        VariantClear(&vDisp); // vDisp.pdispVal->Release();
      CoUninitialize();
    }
  }
protected:
  static Handle<Value> Finalize(const Arguments& args);
  void Finalize();
protected:
  bool finalized;
  static VARIANT vDisp;
};

#endif // __NODE_WIN32OLE_H__
