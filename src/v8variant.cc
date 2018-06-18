/*
  v8variant.cc
*/

#include "v8variant.h"

using namespace v8;
using namespace ole32core;

#define CHECK_OLE_ARGS(args, n, av0, av1) do{ \
    if(args.Length() < n) RETURN_TYPE_ERROR(__FUNCTION__" takes exactly " #n " argument(s)") \
    if(!args[0]->IsString()) RETURN_TYPE_ERROR(__FUNCTION__" the first argument is not a Symbol") \
    if(n == 1) \
      if(args.Length() >= 2) \
        if(!args[1]->IsArray()) RETURN_TYPE_ERROR(__FUNCTION__" the second argument is not an Array") \
        else av1 = args[1]; /* Array */ \
      else av1 = Array::New(isolate, 0); /* change none to Array[] */ \
    else av1 = args[1]; /* may not be Array */ \
    av0 = args[0]; \
  }while(0)

namespace node_win32ole {

    Persistent<FunctionTemplate> V8Variant::clazz;
    Persistent<Function> V8Variant::constructor;

    void V8Variant::Init(Handle<Object> target)
    {
        Isolate* isolate = target->GetIsolate();

        // Prepare constructor template
        Local<String> clazz_name(String::NewFromUtf8(isolate, "V8Variant"));
        Local<FunctionTemplate> clazz = FunctionTemplate::New(isolate, New);
        clazz->SetClassName(clazz_name);
        clazz->InstanceTemplate()->SetInternalFieldCount(2);

        // Prototype
        NODE_SET_PROTOTYPE_METHOD(clazz, "isA", OLEIsA);
        NODE_SET_PROTOTYPE_METHOD(clazz, "vtName", OLEVTName);
        NODE_SET_PROTOTYPE_METHOD(clazz, "toBoolean", OLEBoolean);
        NODE_SET_PROTOTYPE_METHOD(clazz, "toInt32", OLEInt32);
        NODE_SET_PROTOTYPE_METHOD(clazz, "toInt64", OLEInt64);
        NODE_SET_PROTOTYPE_METHOD(clazz, "toNumber", OLENumber);
        NODE_SET_PROTOTYPE_METHOD(clazz, "toDate", OLEDate);
        NODE_SET_PROTOTYPE_METHOD(clazz, "toUtf8", OLEUtf8);
        NODE_SET_PROTOTYPE_METHOD(clazz, "toValue", OLEValue);
        NODE_SET_PROTOTYPE_METHOD(clazz, "call", OLECall);
        NODE_SET_PROTOTYPE_METHOD(clazz, "get", OLEGet);
        NODE_SET_PROTOTYPE_METHOD(clazz, "set", OLESet);
        NODE_SET_PROTOTYPE_METHOD(clazz, "Finalize", Finalize);

        Local<ObjectTemplate> instancetpl = clazz->InstanceTemplate();
        instancetpl->SetCallAsFunctionHandler(OLECallComplete);
        instancetpl->SetNamedPropertyHandler(OLEGetAttr, OLESetAttr);

        constructor.Reset(isolate, clazz->GetFunction());
        target->Set(clazz_name, clazz->GetFunction());
    }

    void V8Variant::New(const FunctionCallbackInfo<Value>& args)
    {
        Isolate* isolate = args.GetIsolate();

        DISPFUNCIN();
        if (!args.IsConstructCall()) RETURN_TYPE_ERROR("Use the new operator to create new V8Variant objects")

        OCVariant *ocv = new OCVariant();
        CHECK_OCV(ocv);

        Local<Object> thisObject = args.This();
        V8Variant *v = new V8Variant(); // must catch exception
        v->Wrap(thisObject); // InternalField[0]
        thisObject->SetInternalField(1, External::New(isolate, ocv));
        //Persistent<Object> objectDisposer = Persistent<Object>::New(thisObject);
        //objectDisposer.MakeWeak(ocv, Dispose);

        args.GetReturnValue().Set(thisObject);
        DISPFUNCOUT();
    }

    void V8Variant::CreateUndefined(Isolate* isolate, Local<Object> &instance)
    {
        DISPFUNCIN();
        Local<Context> context = isolate->GetCurrentContext();
        Local<Function> cons = Local<Function>::New(isolate, constructor);
        instance = cons->NewInstance(context, 0, 0).ToLocalChecked();
        DISPFUNCOUT();
    }

    std::string V8Variant::CreateStdStringMBCSfromUTF8(Handle<Value> v)
    {
        String::Utf8Value u8s(v);
        wchar_t * wcs = u8s2wcs(*u8s);
        if (!wcs) {
            std::cerr << "[Can't allocate string (wcs)]" << std::endl;
            std::cerr.flush();
            return std::string("'!ERROR'");
        }
        char *mbs = wcs2mbs(wcs);
        if (!mbs) {
            free(wcs);
            std::cerr << "[Can't allocate string (mbs)]" << std::endl;
            std::cerr.flush();
            return std::string("'!ERROR'");
        }
        std::string s(mbs);
        free(mbs);
        free(wcs);
        return s;
    }

    OCVariant *V8Variant::CreateOCVariant(Handle<Value> v)
    {
        if (v.IsEmpty()) 
        {
            std::cerr << "[Empty(!)]" << std::endl;
            std::cerr.flush();
        }
        else if (v->IsUndefined())
        {
            std::cerr << "[Undefined(!)]" << std::endl;
            std::cerr.flush();
        }
        else if (v->IsExternal())
        {
            std::cerr << "[External(!)]" << std::endl;
            std::cerr.flush();
        }
        else if (v->IsNativeError())
        {
            std::cerr << "[NativeError(!)]" << std::endl;
            std::cerr.flush();
        }
        else if (v->IsNull()) {
            return new OCVariant();
        }
        else if (v->IsBoolean()) {
            return new OCVariant((bool)(v->BooleanValue() ? !0 : 0));
        }
        else if (v->IsArray()) {
            // VT_BYREF VT_ARRAY VT_SAFEARRAY
            std::cerr << "[Array (not implemented now)]" << std::endl; return NULL;
            std::cerr.flush();
        }
        else if (v->IsInt32()) {
            return new OCVariant((long)v->Int32Value());
#if(0) // may not be supported node.js / v8
        }
        else if (v->IsInt64()) {
            return new OCVariant((long long)v->Int64Value());
#endif
        }
        else if (v->IsNumber()) {
            std::cerr << "[Number (VT_R8 or VT_I8 bug?)]" << std::endl;
            std::cerr.flush();
            // if(v->ToInteger()) =64 is failed ? double : OCVariant((longlong)VT_I8)
            return new OCVariant((double)v->NumberValue()); // double
        }
        else if (v->IsNumberObject()) {
            std::cerr << "[NumberObject (VT_R8 or VT_I8 bug?)]" << std::endl;
            std::cerr.flush();
            // if(v->ToInteger()) =64 is failed ? double : OCVariant((longlong)VT_I8)
            return new OCVariant((double)v->NumberValue()); // double
        }
        else if (v->IsDate()) {
            double d = v->NumberValue();
            time_t sec = (time_t)(d / 1000.0);
            int msec = (int)(d - sec * 1000.0);
            struct tm *t = localtime(&sec); // *** must check locale ***
            if (!t) {
                std::cerr << "[Date may not be valid]" << std::endl;
                std::cerr.flush();
                return NULL;
            }
            SYSTEMTIME syst;
            syst.wYear = t->tm_year + 1900;
            syst.wMonth = t->tm_mon + 1;
            syst.wDay = t->tm_mday;
            syst.wHour = t->tm_hour;
            syst.wMinute = t->tm_min;
            syst.wSecond = t->tm_sec;
            syst.wMilliseconds = msec;
            SystemTimeToVariantTime(&syst, &d);
            return new OCVariant(d, true); // date
        }
        else if (v->IsRegExp()) {
            std::cerr << "[RegExp (bug?)]" << std::endl;
            std::cerr.flush();
            return new OCVariant(CreateStdStringMBCSfromUTF8(v->ToDetailString()));
        }
        else if (v->IsString()) {
            return new OCVariant(CreateStdStringMBCSfromUTF8(v));
        }
        else if (v->IsStringObject()) {
            std::cerr << "[StringObject (bug?)]" << std::endl;
            std::cerr.flush();
            return new OCVariant(CreateStdStringMBCSfromUTF8(v));
        }
        //else if (!v->IsFunction()) {}
        else if (v->IsObject()) {
#if(0)
            std::cerr << "[Object (test)]" << std::endl;
            std::cerr.flush();
#endif
            OCVariant *ocv = castedInternalField<OCVariant>(v->ToObject());
            if (!ocv) {
                std::cerr << "[Object may not be valid (null OCVariant)]" << std::endl;
                std::cerr.flush();
                return NULL;
            }
            // std::cerr << ocv->v.vt;
            return new OCVariant(*ocv);
        }
        else {
            std::cerr << "[unknown type (not implemented now)]" << std::endl;
            std::cerr.flush();
        }
        return new OCVariant(); // NULL
    }

    void V8Variant::OLEIsA(const FunctionCallbackInfo<Value>& args)
    {
        Isolate* isolate = args.GetIsolate();
        DISPFUNCIN();
        OCVariant *ocv = castedInternalField<OCVariant>(args.This());
        CHECK_OCV(ocv);
        args.GetReturnValue().Set(Int32::New(isolate, ocv->v.vt));
        DISPFUNCOUT();
    }

    void V8Variant::OLEVTName(const FunctionCallbackInfo<Value>& args)
    {
        Isolate* isolate = args.GetIsolate();
        DISPFUNCIN();
        OCVariant *ocv = castedInternalField<OCVariant>(args.This());
        CHECK_OCV(ocv);
        Array *a = Array::Cast(*GET_PROP(module_target.Get(isolate), "vt_names"));
        args.GetReturnValue().Set(ARRAY_AT(a, ocv->v.vt));
        DISPFUNCOUT();
    }

    void V8Variant::OLEBoolean(const FunctionCallbackInfo<Value>& args)
    {
        Isolate* isolate = args.GetIsolate();
        DISPFUNCIN();
        OCVariant *ocv = castedInternalField<OCVariant>(args.This());
        CHECK_OCV(ocv);
        if (ocv->v.vt != VT_BOOL) RETURN_TYPE_ERROR("OLEBoolean source type OCVariant is not VT_BOOL")
        bool c_boolVal = ocv->v.boolVal == VARIANT_FALSE ? 0 : !0;
        args.GetReturnValue().Set(Boolean::New(isolate, c_boolVal));
        DISPFUNCOUT();
    }

    void V8Variant::OLEInt32(const FunctionCallbackInfo<Value>& args)
    {
        Isolate* isolate = args.GetIsolate();
        DISPFUNCIN();
        OCVariant *ocv = castedInternalField<OCVariant>(args.This());
        CHECK_OCV(ocv);
        if (ocv->v.vt != VT_I4 && ocv->v.vt != VT_INT
            && ocv->v.vt != VT_UI4 && ocv->v.vt != VT_UINT)
            RETURN_TYPE_ERROR("OLEInt32 source type OCVariant is not VT_I4 nor VT_INT nor VT_UI4 nor VT_UINT")
        DISPFUNCOUT();
        args.GetReturnValue().Set(Int32::New(isolate, ocv->v.lVal));
    }

    void V8Variant::OLEInt64(const FunctionCallbackInfo<Value>& args)
    {
        Isolate* isolate = args.GetIsolate();
        DISPFUNCIN();
        OCVariant *ocv = castedInternalField<OCVariant>(args.This());
        CHECK_OCV(ocv);
        if (ocv->v.vt != VT_I8 && ocv->v.vt != VT_UI8)
            RETURN_TYPE_ERROR("OLEInt64 source type OCVariant is not VT_I8 nor VT_UI8")
        DISPFUNCOUT();
#if(0) // may not be supported node.js / v8
        args.GetReturnValue().Set(Int64::New(isolate, ocv->v.llVal));
#else
        args.GetReturnValue().Set(Number::New(isolate, (double)ocv->v.llVal));
#endif
    }

    void V8Variant::OLENumber(const FunctionCallbackInfo<Value>& args)
    {
        Isolate* isolate = args.GetIsolate();
        DISPFUNCIN();
        OCVariant *ocv = castedInternalField<OCVariant>(args.This());
        CHECK_OCV(ocv);
        if (ocv->v.vt != VT_R8) RETURN_TYPE_ERROR("OLENumber source type OCVariant is not VT_R8")
        DISPFUNCOUT();
        args.GetReturnValue().Set(Number::New(isolate, ocv->v.dblVal));
    }

    void V8Variant::OLEDate(const FunctionCallbackInfo<Value>& args)
    {
        Isolate* isolate = args.GetIsolate();
        DISPFUNCIN();
        OCVariant *ocv = castedInternalField<OCVariant>(args.This());
        CHECK_OCV(ocv);
        if (ocv->v.vt != VT_DATE) RETURN_TYPE_ERROR("OLEDate source type OCVariant is not VT_DATE")
        SYSTEMTIME syst;
        VariantTimeToSystemTime(ocv->v.date, &syst);
        struct tm t = { 0 }; // set t.tm_isdst = 0
        t.tm_year = syst.wYear - 1900;
        t.tm_mon = syst.wMonth - 1;
        t.tm_mday = syst.wDay;
        t.tm_hour = syst.wHour;
        t.tm_min = syst.wMinute;
        t.tm_sec = syst.wSecond;
        DISPFUNCOUT();
        args.GetReturnValue().Set(Date::New(isolate, (double)(mktime(&t) * 1000LL + syst.wMilliseconds)));
    }

    void V8Variant::OLEUtf8(const FunctionCallbackInfo<Value>& args)
    {
        Isolate* isolate = args.GetIsolate();
        DISPFUNCIN();
        OCVariant *ocv = castedInternalField<OCVariant>(args.This());
        CHECK_OCV(ocv);
        if (ocv->v.vt != VT_BSTR) RETURN_TYPE_ERROR("OLEUtf8 source type OCVariant is not VT_BSTR")
        Handle<Value> result;
        if (!ocv->v.bstrVal) result = Undefined(isolate); // or Null(isolate);
        else result = String::NewFromUtf8(isolate, MBCS2UTF8(BSTR2MBCS(ocv->v.bstrVal)).c_str());
        DISPFUNCOUT();
        args.GetReturnValue().Set(result);
    }

    void V8Variant::OLEValue(const FunctionCallbackInfo<Value>& args)
    {
        Isolate* isolate = args.GetIsolate();
        OLETRACEIN();
        OLETRACEVT(args.This());
        OLETRACEFLUSH();
        Local<Object> thisObject = args.This();
        OLE_PROCESS_CARRY_OVER(thisObject);
        OLETRACEVT(thisObject);
        OLETRACEFLUSH();
        OCVariant *ocv = castedInternalField<OCVariant>(thisObject);
        if (!ocv) { std::cerr << "ocv is null"; std::cerr.flush(); }
        CHECK_OCV(ocv);
        switch (ocv->v.vt) {
        case VT_EMPTY: 
            args.GetReturnValue().Set(Undefined(isolate));
            break;
        case VT_NULL:
            args.GetReturnValue().Set(Null(isolate));
            break;
        case VT_DISPATCH: 
            args.GetReturnValue().Set(thisObject); // through it
            break;
        case VT_BOOL:
            OLEBoolean(args); 
            break; 
        case VT_I4:
        case VT_INT:
        case VT_UI4:
        case VT_UINT: 
            OLEInt32(args); 
            break;
        case VT_I8: 
        case VT_UI8:
            OLEInt64(args);
            break;
        case VT_R8: 
            OLENumber(args);
            break;
        case VT_DATE: 
            OLEDate(args);
            break;
        case VT_BSTR: 
            OLEUtf8(args);
            break;
        case VT_ARRAY: 
        case VT_SAFEARRAY:
            std::cerr << "[Array (not implemented now)]" << std::endl;
            std::cerr.flush();
            break;
        default:
            Handle<Value> s = INSTANCE_CALL(thisObject, "vtName", 0, NULL);
            std::cerr << "[unknown type " << ocv->v.vt << ":" << *String::Utf8Value(s);
            std::cerr << " (not implemented now)]" << std::endl;
            std::cerr.flush();
        }
        OLETRACEOUT();
    }

    void V8Variant::OLEFlushCarryOver(Isolate* isolate, Handle<Value> v, Handle<Value> &result)
    {
        OLETRACEIN();
        OLETRACEVT(v->ToObject());
        result = Undefined(isolate);
        V8Variant *v8v = ObjectWrap::Unwrap<V8Variant>(v->ToObject());
        if (v8v->property_carryover.empty()) {
            std::cerr << " *** carryover empty *** " << __FUNCTION__ << std::endl;
            std::cerr.flush();
            // *** or throw exception
        }
        else {
            const char *name = v8v->property_carryover.c_str();
            {
                OLETRACEPREARGV(String::NewSymbol(name));
                OLETRACEARGV();
            }
            OLETRACEFLUSH();
            Handle<Value> argv[] = { String::NewFromUtf8(isolate, name), Array::New(isolate, 0) };
            int argc = sizeof(argv) / sizeof(argv[0]);
            v8v->property_carryover.erase();
            result = INSTANCE_CALL(v->ToObject(), "call", argc, argv);
            if (!result->IsObject()) {
                OCVariant *rv = V8Variant::CreateOCVariant(result);
                CHECK_OCV(rv);
                OCVariant *o = castedInternalField<OCVariant>(v->ToObject());
                CHECK_OCV(o);
                *o = *rv; // copy and don't delete rv
            }
        }
        OLETRACEOUT();
    }

    void V8Variant::OLEInvoke(bool isCall, const FunctionCallbackInfo<Value>& args)
    {
        Isolate* isolate = args.GetIsolate();
        Handle<Value> av0, av1;
        CHECK_OLE_ARGS(args, 1, av0, av1);
        OLETRACEIN();
        OLETRACEVT(args.This());
        OLETRACEARGS();
        OLETRACEFLUSH();
        OCVariant *ocv = castedInternalField<OCVariant>(args.This());
        CHECK_OCV(ocv);
        OCVariant *argchain = NULL;
        Array *a = Array::Cast(*av1);
        for (size_t i = 0; i < a->Length(); ++i) {
            OCVariant *o = V8Variant::CreateOCVariant(
                ARRAY_AT(a, (i ? i : a->Length()) - 1));
            CHECK_OCV(o);
            if (!i) argchain = o;
            else argchain->push(o);
        }
        Local<Object> vResult;
        V8Variant::CreateUndefined(isolate, vResult);
        String::Utf8Value u8s(av0);
        wchar_t *wcs = u8s2wcs(*u8s);
        if (!wcs && argchain) delete argchain;
        bool rcode;
        if (!wcs) {
            std::cerr << "identity is null" << std::endl;
            rcode = false;
        }
        else try {
            OCVariant *rv = isCall // argchain will be deleted automatically
                ? ocv->invoke(wcs, argchain, true)
                : ocv->getProp(wcs, argchain);
            if (rv) {
                OCVariant *o = castedInternalField<OCVariant>(vResult);
                CHECK_OCV(o);
                *o = *rv; // copy and don't delete rv
            }
            rcode = true;
        }
        catch (OLE32coreException e) {
            std::cerr << e.errorMessage(*u8s);
            rcode = false;
        }
        catch (char *e) {
            std::cerr << e << *u8s << std::endl;
            rcode = false;
        }
        free(wcs); // *** it may leak when error ***

        if (rcode) {
            Handle<Value> result = INSTANCE_CALL(vResult, "toValue", 0, NULL);
            args.GetReturnValue().Set(result);
        }
        else {
            THROW_TYPE_ERROR(__FUNCTION__" failed")
        }
        OLETRACEOUT();
    }

    void V8Variant::OLECall(const FunctionCallbackInfo<Value>& args)
    {
        Isolate* isolate = args.GetIsolate();
        OLETRACEIN();
        OLETRACEVT(args.This());
        OLETRACEARGS();
        OLETRACEFLUSH();
        V8Variant::OLEInvoke(true, args); // as Call
        OLETRACEOUT();
    }

    void V8Variant::OLEGet(const FunctionCallbackInfo<Value>& args)
    {
        Isolate* isolate = args.GetIsolate();
        OLETRACEIN();
        OLETRACEVT(args.This());
        OLETRACEARGS();
        OLETRACEFLUSH();
        V8Variant::OLEInvoke(false, args); // as Get
        OLETRACEOUT();
    }

    void V8Variant::OLESet(const FunctionCallbackInfo<Value>& args)
    {
        Isolate* isolate = args.GetIsolate();
        Handle<Value> av0, av1;
        CHECK_OLE_ARGS(args, 2, av0, av1);
        OLETRACEIN();
        OLETRACEVT(args.This());
        OLETRACEARGS();
        OLETRACEFLUSH();
        Local<Object> thisObject = args.This();
        OLE_PROCESS_CARRY_OVER(thisObject);
        OCVariant *ocv = castedInternalField<OCVariant>(thisObject);
        CHECK_OCV(ocv);
        OCVariant *argchain = V8Variant::CreateOCVariant(av1);
        if (!argchain) {
            OLETRACEOUT();
            RETURN_TYPE_ERROR(__FUNCTION__" the second argument is not valid (null OCVariant)");
        }
        bool rcode;
        String::Utf8Value u8s(av0);
        wchar_t *wcs = u8s2wcs(*u8s);
        if (!wcs) {
            std::cerr << "identity is null" << std::endl;
            rcode = false;
        }
        else try {
            ocv->putProp(wcs, argchain); // argchain will be deleted automatically
            rcode = true;
        }
        catch (OLE32coreException e) {
            std::cerr << e.errorMessage(*u8s);
            rcode = false;
        }
        catch (char *e) {
            std::cerr << e << *u8s << std::endl;
            rcode = false;
        }
        free(wcs); // *** it may leak when error ***
        if (!rcode) THROW_ERROR(__FUNCTION__ " failed")
        else args.GetReturnValue().Set(Boolean::New(isolate, true));
        OLETRACEOUT();
    }

    void V8Variant::OLECallComplete(const FunctionCallbackInfo<Value>& args)
    {
        Isolate* isolate = args.GetIsolate();
        OLETRACEIN();
        OLETRACEVT(args.This());
        Handle<Value> result = Undefined(isolate);
        V8Variant *v8v = ObjectWrap::Unwrap<V8Variant>(args.This());
        if (v8v->property_carryover.empty()) {
            std::cerr << " *** carryover empty *** " << __FUNCTION__ << std::endl;
            std::cerr.flush();
            // *** or throw exception
        }
        else {
            const char *name = v8v->property_carryover.c_str();
            {
                OLETRACEPREARGV(String::NewSymbol(name));
                OLETRACEARGV();
            }
            OLETRACEARGS();
            OLETRACEFLUSH();
            Handle<Array> a = Array::New(isolate, args.Length());
            for (int i = 0; i < args.Length(); ++i) ARRAY_SET(a, i, args[i]);
            Handle<Value> argv[] = { String::NewFromUtf8(isolate, name), a };
            int argc = sizeof(argv) / sizeof(argv[0]);
            v8v->property_carryover.erase();
            result = INSTANCE_CALL(args.This(), "call", argc, argv);
        }
        //_
        //Handle<Value> r = INSTANCE_CALL(Handle<Object>::Cast(v), "toValue", 0, NULL);
        OLETRACEOUT();
        args.GetReturnValue().Set(result);
    }

    void V8Variant::OLEGetAttr(Local<String> name, const PropertyCallbackInfo<Value>& args)
    {
        Isolate* isolate = args.GetIsolate();
        Local<Object> thisObject = args.This();
        String::Utf8Value u8name(name);

        OLETRACEIN();
        OLETRACEVT(thisObject);
        {
            OLETRACEPREARGV(name);
            OLETRACEARGV();
        }
        OLETRACEFLUSH();

        // Why GetAttr comes twice for () in the third loop instead of CallComplete ?
        // Because of the Crankshaft v8's run-time optimizer ?
        {
            V8Variant *v8v = ObjectWrap::Unwrap<V8Variant>(thisObject);
            if (!v8v->property_carryover.empty()) {
                if (v8v->property_carryover == *u8name) {
                    OLETRACEOUT();
                    args.GetReturnValue().Set(thisObject); // through it
                    return;
                }
            }
        }

#if(0)
        if (std::string("call") == *u8name || std::string("get") == *u8name
            || std::string("_") == *u8name || std::string("toValue") == *u8name
            //|| std::string("valueOf") == *u8name || std::string("toString") == *u8name
            ) {
            OLE_PROCESS_CARRY_OVER(thisObject);
        }
#else
        if (std::string("set") != *u8name
            && std::string("toBoolean") != *u8name
            && std::string("toInt32") != *u8name && std::string("toInt64") != *u8name
            && std::string("toNumber") != *u8name && std::string("toDate") != *u8name
            && std::string("toUtf8") != *u8name
            && std::string("inspect") != *u8name && std::string("constructor") != *u8name
            && std::string("valueOf") != *u8name && std::string("toString") != *u8name
            && std::string("toLocaleString") != *u8name
            && std::string("hasOwnProperty") != *u8name
            && std::string("isPrototypeOf") != *u8name
            && std::string("propertyIsEnumerable") != *u8name
            //&& std::string("_") != *u8name
            ) 
        {
            OLE_PROCESS_CARRY_OVER(thisObject);
        }
#endif
        OLETRACEVT(thisObject);
        // Can't use INSTANCE_CALL here. (recursion itself)
        // So it returns Object's fundamental function and custom function:
        //   inspect ?, constructor valueOf toString toLocaleString
        //   hasOwnProperty isPrototypeOf propertyIsEnumerable
        static fundamental_attr fundamentals[] = {
          {0, "call", OLECall}, {0, "get", OLEGet}, {0, "set", OLESet},
          {0, "isA", OLEIsA}, {0, "vtName", OLEVTName}, // {"vt_names", ???},
          {!0, "toBoolean", OLEValue},
          {!0, "toInt32", OLEValue}, {!0, "toInt64", OLEValue},
          {!0, "toNumber", OLEValue}, {!0, "toDate", OLEValue},
          {!0, "toUtf8", OLEValue},
          {0, "toValue", OLEValue},
          {0, "inspect", NULL}, {0, "constructor", NULL}, {0, "valueOf", OLEValue},
          {0, "toString", OLEValue}, {0, "toLocaleString", OLEValue},
          {0, "hasOwnProperty", NULL}, {0, "isPrototypeOf", NULL},
          {0, "propertyIsEnumerable", NULL}
        };
        for (int i = 0; i < sizeof(fundamentals) / sizeof(fundamentals[0]); ++i) {
            if (std::string(fundamentals[i].name) != *u8name) continue;
            if (fundamentals[i].obsoleted) {
                std::cerr << " *** ## [." << fundamentals[i].name;
                std::cerr << "()] is obsoleted. ## ***" << std::endl;
                std::cerr.flush();
            }
            OLETRACEFLUSH();
            OLETRACEOUT();
            args.GetReturnValue().Set(FunctionTemplate::New(isolate, 
                fundamentals[i].func, thisObject)->GetFunction());
            return;
        }
        if (std::string("_") == *u8name) { // through it when "_"
#if(0)
            std::cerr << " *** ## [._] is obsoleted. ## ***" << std::endl;
            std::cerr.flush();
#endif
        }
        else {
            Handle<Object> vResult;
            V8Variant::CreateUndefined(isolate, vResult); // uses much memory
            OCVariant *rv = V8Variant::CreateOCVariant(thisObject);
            CHECK_OCV(rv);
            OCVariant *o = castedInternalField<OCVariant>(vResult);
            CHECK_OCV(o);
            *o = *rv; // copy and don't delete rv
            V8Variant *v8v = ObjectWrap::Unwrap<V8Variant>(vResult);
            v8v->property_carryover.assign(*u8name);
            OLETRACEPREARGV(name);
            OLETRACEARGV();
            OLETRACEFLUSH();
            OLETRACEOUT();
            args.GetReturnValue().Set(vResult); // convert and hold it (uses much memory)
            return;
        }
        OLETRACEFLUSH();
        OLETRACEOUT();
        args.GetReturnValue().Set(thisObject); // through it
    }

    void V8Variant::OLESetAttr(Local<String> name, Local<Value> val, const PropertyCallbackInfo<Value>& args)
    {
        Isolate* isolate = args.GetIsolate();
        OLETRACEIN();
        OLETRACEVT(info.This());
        Handle<Value> argv[] = { name, val };
        int argc = sizeof(argv) / sizeof(argv[0]);
        OLETRACEARGV();
        OLETRACEFLUSH();
        Handle<Value> r = INSTANCE_CALL(args.This(), "set", argc, argv);
        args.GetReturnValue().Set(r);
        OLETRACEOUT();
    }

    void V8Variant::Finalize(const FunctionCallbackInfo<Value>& args)
    {
        Isolate* isolate = args.GetIsolate();
        DISPFUNCIN();
#if(0)
        std::cerr << __FUNCTION__ << " Finalizer is called\a" << std::endl;
        std::cerr.flush();
#endif
        Local<Object> thisObject = args.This();
#if(0)
        V8Variant *v = ObjectWrap::Unwrap<V8Variant>(thisObject);
        if (v) delete v; // it has been already deleted ?
        thisObject->SetInternalField(0, External::New(isolate, NULL));
#endif
#if(1) // now GC will call Disposer automatically
        OCVariant *ocv = castedInternalField<OCVariant>(thisObject);
        if (ocv) delete ocv;
#endif
        thisObject->SetInternalField(1, External::New(isolate, NULL));
        args.GetReturnValue().Set(thisObject);
        DISPFUNCOUT();
    }

    void V8Variant::Dispose(Isolate* isolate, Persistent<Object> handle, void *param)
    {
        DISPFUNCIN();
#if(0)
        //  std::cerr << __FUNCTION__ << " Disposer is called\a" << std::endl;
        std::cerr << __FUNCTION__ << " Disposer is called" << std::endl;
        std::cerr.flush();
#endif
        Local<Object> thisObject = handle.Get(isolate);
#if(0) // it has been already deleted ?
        V8Variant *v = ObjectWrap::Unwrap<V8Variant>(thisObject);
        if (!v) {
            std::cerr << __FUNCTION__;
            std::cerr << "InternalField[0] has been already deleted" << std::endl;
            std::cerr.flush();
        }
        else delete v; // it has been already deleted ?
        BEVERIFY(done, thisObject->InternalFieldCount() > 0);
        thisObject->SetInternalField(0, External::New(isolate, NULL));
#endif
        OCVariant *p = castedInternalField<OCVariant>(thisObject);
        if (!p) {
            std::cerr << __FUNCTION__;
            std::cerr << "InternalField[1] has been already deleted" << std::endl;
            std::cerr.flush();
        }
        //  else{
        OCVariant *ocv = static_cast<OCVariant *>(param); // ocv may be same as p
        if (ocv) delete ocv;
        //  }
        BEVERIFY(done, thisObject->InternalFieldCount() > 1);
        thisObject->SetInternalField(1, External::New(isolate, NULL));
    done:
        handle.Reset();
        DISPFUNCOUT();
    }

    void V8Variant::Finalize()
    {
        assert(!finalized);
        finalized = true;
    }

} // namespace node_win32ole
