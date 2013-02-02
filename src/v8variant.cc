/*
  v8variant.cc
*/

#include "v8variant.h"

using namespace v8;
using namespace ole32core;

namespace node_win32ole {

Persistent<FunctionTemplate> V8Variant::constructor_template;

void V8Variant::Init(Handle<Object> target)
{
  HandleScope scope;
  Local<FunctionTemplate> t = FunctionTemplate::New(New);
  constructor_template = Persistent<FunctionTemplate>::New(t);
  constructor_template->InstanceTemplate()->SetInternalFieldCount(1);
  constructor_template->SetClassName(String::NewSymbol("V8Variant"));
  NODE_SET_PROTOTYPE_METHOD(constructor_template, "Del", Del);
  target->Set(String::NewSymbol("V8Variant"),
    constructor_template->GetFunction());
}

Handle<Value> V8Variant::New(const Arguments& args)
{
  HandleScope scope;
  std::cerr << "-called " << __FUNCTION__ << std::endl;
  if(!args.IsConstructCall())
    return ThrowException(Exception::TypeError(
      String::New("Use the new operator to create new V8Variant objects")));
  return args.This();
}

Handle<Value> V8Variant::Del(const Arguments& args)
{
  HandleScope scope;
  std::cerr << "-called " << __FUNCTION__ << std::endl;
  return args.This();
}

void V8Variant::Finalize()
{
  assert(!finalized);
  finalized = true;
}

} // namespace node_win32ole
