/*
  win32ole_gettimeofday.cc
*/

#include <node_buffer.h>
#include <node.h>
#include <v8.h>

#ifdef _WIN32
#include <sys/timeb.h>
#endif

#define XASSERT(e) do{ \
  if(!(e)){ \
    fprintf(stderr, "ERROR: file %s (%d): ", __FILE__, __LINE__); \
    fprintf(stderr, "%s\n", #e); \
    goto XASSERTSKIP; \
  } \
}while(0)

using namespace v8;

Handle<Value> Method_gettimeofday(const Arguments& args)
{
  boolean result = false;
  HandleScope scope;
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
    XASSERT(args[0]->IsObject());
    v8::Handle<v8::Object> buf = args[0]->ToObject();
    XASSERT(node::Buffer::Length(buf) == sizeof(struct timeval));
    memcpy(node::Buffer::Data(buf), &tv, sizeof(struct timeval));
    result = true;
XASSERTSKIP:;
  }
  return scope.Close(Boolean::New(result));
}
