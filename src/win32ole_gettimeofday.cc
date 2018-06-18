/*
  win32ole_gettimeofday.cc
*/

#include "node_win32ole.h"
#include "ole32core.h"

#ifdef _WIN32
#include <sys/timeb.h>
#endif

using namespace v8;
using namespace ole32core;

namespace node_win32ole {

    void timeval_gettimeofday(struct timeval *ptv)
    {
#ifdef _WIN32
        _timeb btv;
        _ftime(&btv);
        ptv->tv_sec = (long)btv.time; // unsafe convert __time64_t to long
        ptv->tv_usec = 1000LL * btv.millitm;
#else
        gettimeofday(ptv, NULL); // no &tz
#endif
    }

    void Method_gettimeofday(const FunctionCallbackInfo<Value>& args)
    {
        Isolate* isolate = args.GetIsolate();
        boolean result = false;
        if BVERIFY(args.Length() >= 2) {
            struct timeval tv;
            timeval_gettimeofday(&tv);
            if BVERIFY(args[0]->IsObject()) {
                v8::Handle<v8::Object> buf = args[0]->ToObject();
                if BVERIFY(node::Buffer::Length(buf) == sizeof(struct timeval)) {
                    memcpy(node::Buffer::Data(buf), &tv, sizeof(struct timeval));
                    result = true;
                }
            }
        }
        args.GetReturnValue().Set(Boolean::New(isolate, result));
    }

    void Method_sleep(const FunctionCallbackInfo<Value>& args) // ms, bool: msg, bool: \n
    {
        Isolate* isolate = args.GetIsolate();
        if (!BVERIFY(args.Length() >= 1)) {
            args.GetReturnValue().Set(Boolean::New(isolate, false));
            return;
        }
        if (!args[0]->IsInt32()) RETURN_TYPE_ERROR("type of argument 1 must be Int32");
        long ms = args[0]->Int32Value();
        bool msg = false;
        if (args.Length() >= 2) {
            if (!args[1]->IsBoolean()) RETURN_TYPE_ERROR("type of argument 2 must be Boolean")
            msg = args[1]->BooleanValue();
        }
        bool crlf = false;
        if (args.Length() >= 3) {
            if (!args[2]->IsBoolean()) RETURN_TYPE_ERROR("type of argument 3 must be Boolean")
            crlf = args[2]->BooleanValue();
        }
        if (ms) {
            if (msg) {
                printf("waiting %ld milliseconds...", ms);
                if (crlf) printf("\n");
            }
            struct timeval tv_start;
            timeval_gettimeofday(&tv_start);
            while (true) {
                struct timeval tv_now;
                timeval_gettimeofday(&tv_now);
                // it is wrong way (must release CPU)
                if ((double)tv_now.tv_sec + (double)tv_now.tv_usec / 1000. / 1000.
                    >= (double)tv_start.tv_sec + (double)tv_start.tv_usec / 1000. / 1000.
                    + (double)ms / 1000.) break;
            }
        }
        args.GetReturnValue().Set(Boolean::New(isolate, true));
    }

} // namespace node_win32ole
