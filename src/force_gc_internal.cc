/*
  force_gc_internal.cc
*/

// #define __OLD_FORCE_GC_INTERNAL__
#ifdef __OLD_FORCE_GC_INTERNAL__
#include <v8.h> // position: v8/include/v8.h
#include <../src/v8.h> // position: v8/src/v8.h
// *** don't include other headers ***
#else
#include "node_win32ole.h"
#endif
#include <iostream>

using namespace v8;

namespace node_win32ole {

Handle<Value> Method_force_gc_internal(const Arguments& args) // v8/src/v8.h
{
  HandleScope scope;
  std::cerr << "-in: " __FUNCTION__ << std::endl;
  if(args.Length() < 1)
    return ThrowException(Exception::TypeError(
      String::New("this function takes at least 1 argument(s)")));
  if(!args[0]->IsInt32())
    return ThrowException(Exception::TypeError(
      String::New("type of argument 1 must be Int32")));
  int flags = (int)args[0]->Int32Value();
#ifdef __OLD_FORCE_GC_INTERNAL__
  std::string reason("called from win32ole.force_gc_internal"); // default
  if(args.Length() >= 2){
    if(!args[1]->IsString())
      return ThrowException(Exception::TypeError(
        String::New("type of argument 2 must be (utf8)String")));
    String::Utf8Value u8s(args[1]);
    reason = std::string(*u8s);
  }
/*
  ( v8 bundled node 0.8.18 )
  0: kNoGCFlags
  1: kSweepPreciselyMask
  2: kReduceMemoryFootprintMask
  4: kAbortIncrementalMarkingMask
  kMakeHeapIterableMask = kSweepPreciselyMask | kAbortIncrementalMarkingMask
*/
  // Performs a full garbage collection. If (flags & kMakeHeapIterableMask) is
  // non-zero, then the slower precise sweeper is used, which leaves the heap
  // in a state where we can iterate over the heap visiting all objects.
  v8::internal::Heap::CollectAllGarbage(flags, reason.c_str());
  // Last hope GC, should try to squeeze as much as possible.
  // v8::internal::Heap::CollectAllAvailableGarbage(reason.c_str());
#else
  while(!v8::V8::IdleNotification()){}
#endif
  std::cerr << "-out: " __FUNCTION__ << std::endl;
  return scope.Close(Boolean::New(true));
}

} // namespace node_win32ole
