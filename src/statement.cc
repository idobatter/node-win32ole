/*
  statement.cc
*/

#include "statement.h"
#include "ole32core.h"
#include "v8variant.h"

using namespace v8;
using namespace ole32core;

namespace node_win32ole {

Persistent<FunctionTemplate> Statement::clazz;

void Statement::Init(Handle<Object> target)
{
  HandleScope scope;
  Local<FunctionTemplate> t = FunctionTemplate::New(New);
  clazz = Persistent<FunctionTemplate>::New(t);
  clazz->InstanceTemplate()->SetInternalFieldCount(1);
  clazz->SetClassName(String::NewSymbol("Statement"));
  NODE_SET_PROTOTYPE_METHOD(clazz, "Dispatch", Dispatch);
  NODE_SET_PROTOTYPE_METHOD(clazz, "Finalize", Finalize);
  target->Set(String::NewSymbol("Statement"), clazz->GetFunction());
}

Handle<Value> Statement::New(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
  if(!args.IsConstructCall())
    return ThrowException(Exception::TypeError(
      String::New("Use the new operator to create new Statement objects")));
  OLE32core *oc = new OLE32core();
  if(!oc)
    return ThrowException(Exception::TypeError(
      String::New("Can't create new Statement object (null OLE32core)")));
  Local<Object> thisObject = args.This();
  thisObject->SetInternalField(0, External::New(oc));
  Persistent<Object> objectDisposer = Persistent<Object>::New(thisObject);
  objectDisposer.MakeWeak(oc, Dispose);
  DISPFUNCOUT();
  return args.This();
}

Handle<Value> Statement::Dispatch(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
  boolean result = false;
  BEVERIFY(done, args.Length() >= 2);
  BEVERIFY(done, args[0]->IsString());
  BEVERIFY(done, args[1]->IsString());
  char cstr_locale[256];
  {
    String::Utf8Value u8s_locale(args[1]);
    strncpy(cstr_locale, *u8s_locale, sizeof(cstr_locale));
  }
  wchar_t *wcs;
  {
    String::Utf8Value u8s(args[0]); // must create here
    wcs = u8s2wcs(*u8s);
  }
  BEVERIFY(done, wcs);
#ifdef DEBUG
  char *mbs = wcs2mbs(wcs);
  if(!mbs) free(wcs);
  BEVERIFY(done, mbs);
  fprintf(stderr, "ProgID: %s\n", mbs);
  free(mbs);
#endif
  CLSID clsid;
  HRESULT hr = CLSIDFromProgID(wcs, &clsid);
  free(wcs);
  BEVERIFY(done, !FAILED(hr));
#ifdef DEBUG
  fprintf(stderr, "clsid:"); // 00024500-0000-0000-c000-000000000046 (Excel) ok
  for(int i = 0; i < sizeof(CLSID); ++i)
    fprintf(stderr, " %02x", ((unsigned char *)&clsid)[i]);
  fprintf(stderr, "\n");
#endif
  OLE32core *oc = castedInternalField<OLE32core>(args.This());
  if(!oc)
    return ThrowException(Exception::TypeError(
      String::New("Can't access to Statement object (null OLE32core)")));
  BEVERIFY(done, oc->connect(cstr_locale));
  Handle<Object> vApp = V8Variant::CreateUndefined();
  BEVERIFY(done, !vApp.IsEmpty());
  BEVERIFY(done, !vApp->IsUndefined());
  BEVERIFY(done, vApp->IsObject());
  OCVariant *app = castedInternalField<OCVariant>(vApp);
  if(!app)
    return ThrowException(Exception::TypeError(
      String::New("Can't access to V8Variant object (null OCVariant)")));
  app->v.vt = VT_DISPATCH;
  // When 'CoInitialize(NULL)' is not called first (and on the same instance),
  // next functions will return many errors.
  // (old style) GetActiveObject() returns 0x000036b7
  //   The requested lookup key was not found in any active activation context.
  // (OLE2) CoCreateInstance() returns 0x000003f0
  //   An attempt was made to reference a token that does not exist.
#ifdef DEBUG // obsolete (it needs that OLE target has been already executed)
  IUnknown *pUnk;
  hr = GetActiveObject(clsid, NULL, (IUnknown **)&pUnk);
  BEVERIFY(done, !FAILED(hr));
  hr = pUnk->QueryInterface(IID_IDispatch, (void **)&app->v.pdispVal);
  pUnk->Release();
#else
  // C -> C++ changes types (&clsid -> clsid, &IID_IDispatch -> IID_IDispatch)
  hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER|CLSCTX_LOCAL_SERVER,
    IID_IDispatch, (void **)&app->v.pdispVal);
#endif
  BEVERIFY(done, !FAILED(hr));
  return scope.Close(vApp);

  try{
    app->putProp(L"Visible", new OCVariant((long)1));
    OCVariant *books = app->getProp(L"Workbooks");
    OCVariant *book = books->invoke(L"Add", NULL, true);
    OCVariant *sheet = book->getProp(L"Worksheets", new OCVariant((long)1));
    sheet->putProp(L"Name", new OCVariant("sheetnameA mbs"));
    {
      OCVariant *argchain = new OCVariant((long)2);
      argchain->push(new OCVariant((long)1));
      OCVariant *cells = sheet->getProp(L"Cells", argchain);
      if(!cells){
        std::cerr << " cells not assigned\n";
      }else{
        cells->putProp(L"Value", new OCVariant("test mbs"));
        delete cells;
      }
    }
    {
      OCVariant *argchain1 = new OCVariant((long)2);
      argchain1->push(new OCVariant((long)2));
      OCVariant *cells0 = sheet->getProp(L"Cells", argchain1);
      if(!cells0){
        std::cerr << " cells0 not assigned\n";
      }else{
        OCVariant *argchain2 = new OCVariant((long)4);
        argchain2->push(new OCVariant((long)4));
        OCVariant *cells1 = sheet->getProp(L"Cells", argchain2);
        if(!cells1){
          std::cerr << " cells1 not assigned\n";
        }else{
          OCVariant *argchain0 = new OCVariant(*cells1); // need copy
          argchain0->push(new OCVariant(*cells0)); // need copy
          OCVariant *rg = sheet->getProp(L"Range", argchain0);
          if(!rg){
            std::cerr << " rg not assigned\n";
          }else{
            rg->putProp(L"RowHeight", new OCVariant(5.18));
            rg->putProp(L"ColumnWidth", new OCVariant(0.58));
            OCVariant *interior = rg->getProp(L"Interior");
            if(!interior){
              std::cerr << " rg.Interior not assigned\n";
            }else{
              interior->putProp(L"ColorIndex", new OCVariant((long)6));
              delete interior;
            }
            delete rg;
          }
          delete cells1;
        }
        delete cells0;
      }
    }

Handle<Value> v = module_target->Get(String::NewSymbol("MODULEDIRNAME"));
BEVERIFY(done, !v.IsEmpty());
BEVERIFY(done, !v->IsUndefined());
BEVERIFY(done, !v->IsObject());
BEVERIFY(done, v->IsString());
String::Utf8Value s(v);
std::cerr << *s << std::endl;
std::string outfile(std::string(*s) + "\\..\\test\\tmp\\testfilembs.xls");

    book->invoke(L"SaveAs", new OCVariant(outfile));
    std::cerr << "saved to: " << outfile << std::endl;
    app->putProp(L"ScreenUpdating", new OCVariant((long)1));
    books->invoke(L"Close");
    app->invoke(L"Quit");
    delete sheet;
    delete book;
    delete books;
  }catch(OLE32coreException e){ std::cerr << e.errorMessage("all"); goto done;
  }catch(char *e){ std::cerr << e << "[all]" << std::endl; goto done;
  }
  // app will be deleted in the destructor of xApp
  result = true;
done:
  DISPFUNCOUT();
  return scope.Close(Boolean::New(result));
}

Handle<Value> Statement::Finalize(const Arguments& args)
{
  HandleScope scope;
  DISPFUNCIN();
  Local<Object> thisObject = args.This();
#if(1) // now GC will call Disposer automatically
  OLE32core *oc = castedInternalField<OLE32core>(thisObject);
  if(oc){ oc->disconnect(); delete oc; }
#endif
  thisObject->SetInternalField(0, External::New(NULL));
#if(0) // to force GC (shold not use at here ?)
  // v8::internal::Heap::CollectAllGarbage(); // obsolete
  while(!v8::V8::IdleNotification());
#endif
  DISPFUNCOUT();
  return args.This();
}

void Statement::Dispose(Persistent<Value> handle, void *param)
{
  DISPFUNCIN();
  OLE32core *oc = static_cast<OLE32core *>(param);
  if(oc){ oc->disconnect(); delete oc; }
  // thisObject->SetInternalField(0, External::New(NULL));
  handle.Dispose();
  DISPFUNCOUT();
}

void Statement::Finalize()
{
  assert(!finalized);
  finalized = true;
}

} // namespace node_win32ole
