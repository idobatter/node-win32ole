#ifndef __NODE_WIN32OLE_H__
#define __NODE_WIN32OLE_H__

#include <node_buffer.h>
#include <node.h>
#include <v8.h>

using namespace v8;

#define BASSERT(x) chkerr((BOOL)(x), __FILE__, __LINE__, __FUNCTION__, #x)
#define BVERIFY(x) BASSERT(x)
#if defined(_DEBUG) || defined(DEBUG)
#define DASSERT(x) BASSERT(x)
#define DVERIFY(x) DASSERT(x)
#else // RELEASE
#define DASSERT(x)
#define DVERIFY(x) (x)
#endif
extern BOOL chkerr(BOOL b, char *m, int n, char *f, char *e);
#define BEVERIFY(y, x) if(!BVERIFY(x)){ goto y; }
#define DEVERIFY(y, x) if(!DVERIFY(x)){ goto y; }

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
  static VARIANT vDisp;
};

#endif // __NODE_WIN32OLE_H__
