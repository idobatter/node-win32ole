/*
  win32ole_gettimeofday.cc
*/

#include "node_win32ole.h"

#ifdef _WIN32
#include <sys/timeb.h>
#endif

using namespace v8;

Handle<Value> Method_gettimeofday(const Arguments& args)
{
  HandleScope scope;
  boolean result = false;
  if(args.Length() >= 2){
    struct timeval tv;
#ifdef _WIN32
    _timeb btv;
    _ftime(&btv);
    tv.tv_sec = btv.time; // unsafe convert __time64_t to long
    tv.tv_usec = 1000LL * btv.millitm;
#else
    gettimeofday(&tv, NULL); // no &tz
#endif
    BEVERIFY(done, args[0]->IsObject());
    v8::Handle<v8::Object> buf = args[0]->ToObject();
    BEVERIFY(done, node::Buffer::Length(buf) == sizeof(struct timeval));
    memcpy(node::Buffer::Data(buf), &tv, sizeof(struct timeval));
    result = true;
done:
    ;
  }
  return scope.Close(Boolean::New(result));
}
