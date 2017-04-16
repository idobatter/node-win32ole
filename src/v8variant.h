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
  void (*func)(const FunctionCallbackInfo<Value>& args);
} fundamental_attr;

class V8Variant : public node::ObjectWrap {
public:
	static Persistent<FunctionTemplate> clazz;
	static void Init(Handle<Object> target);
	static void New(const FunctionCallbackInfo<Value>& args);

	static std::string CreateStdStringMBCSfromUTF8(Handle<Value> v); // *** p.
	static OCVariant *CreateOCVariant(Handle<Value> v); // *** private
	static void CreateUndefined(Isolate* isolate, Local<Object> &instance); // *** private
	static void OLEIsA(const FunctionCallbackInfo<Value>& args);
	static void OLEVTName(const FunctionCallbackInfo<Value>& args);
	static void OLEBoolean(const FunctionCallbackInfo<Value>& args); // *** p.
	static void OLEInt32(const FunctionCallbackInfo<Value>& args); // *** p.
	static void OLEInt64(const FunctionCallbackInfo<Value>& args); // *** p.
	static void OLENumber(const FunctionCallbackInfo<Value>& args); // *** p.
	static void OLEDate(const FunctionCallbackInfo<Value>& args); // *** p.
	static void OLEUtf8(const FunctionCallbackInfo<Value>& args); // *** p.
	static void OLEValue(const FunctionCallbackInfo<Value>& args);
	static void OLEFlushCarryOver(Isolate* isolate, Handle<Value> v, Handle<Value> &result); // *** p.
	static void OLEInvoke(bool isCall, const FunctionCallbackInfo<Value>& args); // *** p.
	static void OLECall(const FunctionCallbackInfo<Value>& args);
	static void OLEGet(const FunctionCallbackInfo<Value>& args);
	static void OLESet(const FunctionCallbackInfo<Value>& args);
	static void OLECallComplete(const FunctionCallbackInfo<Value>& args); // *** p.
	static void OLEGetAttr(Local<String> name, const PropertyCallbackInfo<Value>& args); // *** p.
	static void OLESetAttr(Local<String> name, Local<Value> val, const PropertyCallbackInfo<Value>& args); // *** p.
	static void Finalize(const FunctionCallbackInfo<Value>& args);
public:
	V8Variant() : node::ObjectWrap(), finalized(false), property_carryover() {}
	~V8Variant() { if(!finalized) Finalize(); }
protected:
	static void Dispose(Isolate* isolate, Persistent<Object> handle, void *param);
	void Finalize();
protected:
	bool finalized;
	static Persistent<Function> constructor;
	std::string property_carryover;
};

} // namespace node_win32ole

#endif // __V8VARIANT_H__
