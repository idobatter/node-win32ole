/*
  node_win32ole.cc
  $ cp ~/.node-gyp/0.8.18/ia32/node.lib to ~/.node-gyp/0.8.18/node.lib
  $ node-gyp configure
  $ node-gyp build
  $ node test/init_win32ole.test.js
*/

#include <node.h>
#include <v8.h>

using namespace v8;

Handle<Value> Method_gettimeofday(const Arguments& args);

Handle<Value> Method_print(const Arguments& args)
{
  HandleScope scope;
  if(args.Length() >= 1){
    String::Utf8Value s(args[0]);
    char *cs = *s;
    printf("%s\n", cs);
  }
  return scope.Close(Boolean::New(true));
}

void init(Handle<Object> target)
{
  target->Set(String::NewSymbol("print"),
    FunctionTemplate::New(Method_print)->GetFunction());
  target->Set(String::NewSymbol("gettimeofday"),
    FunctionTemplate::New(Method_gettimeofday)->GetFunction());
}

NODE_MODULE(node_win32ole, init)
