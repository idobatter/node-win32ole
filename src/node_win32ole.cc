/*
  node_win32ole.cc
  $ cp ~/.node-gyp/0.8.18/ia32/node.lib to ~/.node-gyp/0.8.18/node.lib
  $ node-gyp configure
  $ node-gyp build
  $ node test/init_win32ole.test.js
*/

#include "node_win32ole.h"

using namespace v8;

Persistent<Object> module_target;
Persistent<FunctionTemplate> Statement::constructor_template;
OLE32core Statement::oc;

void Statement::Init(Handle<Object> target)
{
  HandleScope scope;
  Local<FunctionTemplate> t = FunctionTemplate::New(New);
  constructor_template = Persistent<FunctionTemplate>::New(t);
  constructor_template->InstanceTemplate()->SetInternalFieldCount(1);
  constructor_template->SetClassName(String::NewSymbol("Statement"));
  NODE_SET_PROTOTYPE_METHOD(constructor_template, "Dispatch", Dispatch);
  NODE_SET_PROTOTYPE_METHOD(constructor_template, "Finalize", Finalize);
  target->Set(String::NewSymbol("Statement"),
    constructor_template->GetFunction());
}

Handle<Value> Statement::New(const Arguments& args)
{
  HandleScope scope;
  if(!args.IsConstructCall())
    return ThrowException(Exception::TypeError(
      String::New("Use the new operator to create new Statement objects")));
  return args.This();
}

Handle<Value> Statement::Dispatch(const Arguments& args)
{
  HandleScope scope;
  boolean result = false;

#if(0) // ---- testing ----

  if(!args[0]->IsString()){
    fprintf(stderr, "error: args[0] is not a string");
    return scope.Close(Boolean::New(result));
  }
  String::Utf8Value s(args[0]);
  wchar_t *wcs = u8s2wcs(*s);
  BEVERIFY(done, wcs);
#ifdef DEBUG
  char *mbs = wcs2mbs(wcs);
  if(!mbs) free(wcs);
  BEVERIFY(done, mbs);
  fprintf(stderr, "ProgID: %s\n", mbs); // Excel.Application
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
  // When this function is not called first (and on the same instance),
  // next functions will return many errors.
  CoInitialize(NULL);
  // (old style) GetActiveObject() returns 0x000036b7
  //   The requested lookup key was not found in any active activation context.
  // (OLE2) CoCreateInstance() returns 0x000003f0
  //   An attempt was made to reference a token that does not exist.
#ifdef DEBUG // old style (needs pre executed)
  IUnknown *pUnk;
  hr = GetActiveObject(clsid, NULL, (IUnknown **)&pUnk);
  BEVERIFY(done, !FAILED(hr));
  hr = pUnk->QueryInterface(IID_IDispatch, (void **)&vDisp.pdispVal);
  BEVERIFY(done, !FAILED(hr));
  pUnk->Release();
#else
  // C -> C++ changes types (&clsid -> clsid, &IID_IDispatch -> IID_IDispatch)
  vDisp.vt = VT_DISPATCH;
  hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER|CLSCTX_LOCAL_SERVER,
    IID_IDispatch, (void **)&vDisp.pdispVal);
  BEVERIFY(done, !FAILED(hr));
#endif

#else // ---- testing ----

  std::cerr << "-in\n";
  try{
    OCVariant *app = oc.connect("Japanese", true);
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
    std::string outfile("c:\\prj\\node-win32ole\\test\\tmp\\testfilembs.xls");
    book->invoke(L"SaveAs", new OCVariant(outfile));
    std::cerr << "saved to: " << outfile << std::endl;
    app->putProp(L"ScreenUpdating", new OCVariant((long)1));
    books->invoke(L"Close");
    app->invoke(L"Quit");
    delete sheet;
    delete book;
    delete books;
    delete app;
  }catch(OLE32coreException e){ std::cerr << e.errorMessage("all"); goto done;
  }catch(char *e){ std::cerr << e << "[all]" << std::endl; goto done;
  }
  std::cerr << "-out\n";

#endif // ---- testing ----

  result = true;
done:
  return scope.Close(Boolean::New(result));
}

Handle<Value> Statement::Finalize(const Arguments& args)
{
  HandleScope scope;
#if(0) // ---- testing ----
  if(vDisp.vt == VT_DISPATCH && vDisp.pdispVal)
    VariantClear(&vDisp); // vDisp.pdispVal->Release();
  CoUninitialize();
#else // ---- testing ----
  oc.disconnect();
#endif // ---- testing ----
  return args.This();
}

void Statement::Finalize()
{
  assert(!finalized);
  finalized = true;
}

Handle<Value> Method_version(const Arguments& args)
{
  HandleScope scope;
  if(args.Length() >= 0){
    //
  }
#ifdef DEBUG
  node::node_module_struct *nms = node::get_builtin_module("Array");
  if(!nms) return scope.Close(String::NewSymbol("not found"));
  return scope.Close(String::New(nms->modname));
#else
  return scope.Close(module_target->Get(String::NewSymbol("VERSION")));
#endif
}

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
  module_target = Persistent<Object>::New(target);
  Statement::Init(target);
  target->Set(String::NewSymbol("VERSION"),
    String::New("0.0.0 (will be set later)"),
    static_cast<PropertyAttribute>(DontDelete));
  target->Set(String::NewSymbol("SOURCE_TIMESTAMP"),
    String::NewSymbol(__FILE__ " " __DATE__ " " __TIME__),
    static_cast<PropertyAttribute>(ReadOnly | DontDelete));
  target->Set(String::NewSymbol("version"),
    FunctionTemplate::New(Method_version)->GetFunction());
  target->Set(String::NewSymbol("print"),
    FunctionTemplate::New(Method_print)->GetFunction());
  target->Set(String::NewSymbol("gettimeofday"),
    FunctionTemplate::New(Method_gettimeofday)->GetFunction());
}

NODE_MODULE(node_win32ole, init)
