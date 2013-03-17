# NAME

node-win32ole - Asynchronous, non-blocking win32ole bindings for [node.js](https://github.com/joyent/node) powered by v8 engine .

win32ole makes accessibility from node.js to Excel, Word, Access, Outlook, InternetExplorer, WSH ( ActiveXObject / COM ) and so on. It does not need TypeLibrary.


# USAGE

Install with `npm install win32ole`.

It works as... (version 0.1.x)

``` js
try{
  var win32ole = require('win32ole');
  // var xl = new ActiveXObject('Excel.Application'); // You may write it as:
  var xl = win32ole.client.Dispatch('Excel.Application');
  xl.Visible = true;
  var book = xl.Workbooks.Add();
  var sheet = book.Worksheets(1);
  try{
    sheet.Name = 'sheetnameA utf8';
    sheet.Cells(1, 2).Value = 'test utf8';
    var rg = sheet.Range(sheet.Cells(2, 2), sheet.Cells(4, 4));
    rg.RowHeight = 5.18;
    rg.ColumnWidth = 0.58;
    rg.Interior.ColorIndex = 6; // Yellow
    var result = book.SaveAs('testfileutf8.xls');
    console.log(result);
  }catch(e){
    console.log('(exception cached)\n' + e);
  }
  xl.ScreenUpdating = true;
  xl.Workbooks.Close();
  xl.Quit();
}catch(e){
  console.log('*** exception cached ***\n' + e);
}
```

There are 3 ways to make force Garbage Collection for node.js / v8 .

- 1. use huge memory to run GC automatically ( causes abnormal termination )
- 2. win32ole.force_gc_extension(1);
- 3. win32ole.force_gc_internal(1);

see also [examples/ole_args_test_client.js](https://github.com/idobatter/node-win32ole/blob/master/examples/ole_args_test_client.js)


# Tutorial and Examples

- [test/init_win32ole.test.js](https://github.com/idobatter/node-win32ole/blob/master/test/init_win32ole.test.js)
- [test/unicode.test.js](https://github.com/idobatter/node-win32ole/blob/master/test/unicode.test.js)
- [examples/maze_creator.js](https://github.com/idobatter/node-win32ole/blob/master/examples/maze_creator.js)
- [examples/maze_solver.js](https://github.com/idobatter/node-win32ole/blob/master/examples/maze_solver.js)
- [examples/word_sample.js](https://github.com/idobatter/node-win32ole/blob/master/examples/word_sample.js)
- [examples/access_mdb_sample.js](https://github.com/idobatter/node-win32ole/blob/master/examples/access_mdb_sample.js)
- [examples/outlook_sample.js](https://github.com/idobatter/node-win32ole/blob/master/examples/outlook_sample.js)
- [examples/ie_sample.js](https://github.com/idobatter/node-win32ole/blob/master/examples/ie_sample.js)
- [examples/typelibrary_sample.js](https://github.com/idobatter/node-win32ole/blob/master/examples/typelibrary_sample.js)
- [examples/uncfinder_sample.js](https://github.com/idobatter/node-win32ole/blob/master/examples/uncfinder_sample.js)
- [examples/activex_filesystemobject_sample.js](https://github.com/idobatter/node-win32ole/blob/master/examples/activex_filesystemobject_sample.js)
- [examples/wmi_sample.js](https://github.com/idobatter/node-win32ole/blob/master/examples/wmi_sample.js)
- [examples/wsh_sample.js](https://github.com/idobatter/node-win32ole/blob/master/examples/wsh_sample.js)
- [examples/ole_args_test_client.js](https://github.com/idobatter/node-win32ole/blob/master/examples/ole_args_test_client.js)
- [examples/ole_args_test_client_metamorphoses.js](https://github.com/idobatter/node-win32ole/blob/master/examples/ole_args_test_client_metamorphoses.js)


# Other built in functions

* win32ole.version(void) // returns version string
* win32ole.printACP(utf8string) // Utf8 to .ACP
* win32ole.print(utf8string) // ASCII
* win32ole.gettimeofday(struct timeval &tv, null) // now arg2 is not used
* win32ole.sleep(long milliseconds, bool withmessage=false, bool with\n=false)
* win32ole.force_gc_extension(long flag) // now flag is dummy
* win32ole.force_gc_internal(long flag, string) // now flag is dummy


# FEATURES

* fix BUG: date
* BUG: A few samples in win32ole@0.1.0 needs '._' ideom.
* When you use unary operator '!' at the place that needs boolean CONDITION (for example 'while(!obj.status){...}') , you must write 'while(!obj.status._){...}' to complete v8::Object::ToBoolean() conversion. (NamedPropertyHandler will not be called because v8::Object::ToBoolean() is called directly for unary operator '!' instead of v8::Object::valueOf() in ParseUnaryExpression() v8/src/parser.cc .) Do you know how to fake it?
* V8Variant::OLEGetAttr returns a copy of object, so it uses much memory. I want to fix it.
* Now '._' ideom is obsoleted.
* Remove 'node-proxy' from dependencies list.
* Change default branch to dev0.1.0 .
* BUG: Some samples in between win32ole@0.0.25 and win32ole@0.0.28 ( examples/maze_creator.js examples/maze_solver.js ) uses huge memory and many disposers will run by v8 GC when maze size is 20*30. I think that each encapsulated V8Variant (by node-proxy) may be big object. So I will try to use v8 accessor handlers ( SetCallAsFunctionHandler / SetNamedPropertyHandler / SetIndexedPropertyHandler ) instead of ( '__noSuchMethod__' / '__noSuchGetter__' / '__noSuchSetter__' ) by node-proxy.
* So much implements. (can not handle some COM VARIANT types, array etc.)
* Bug fix. (throws exception when failed to Invoke(), and many test message.)
* Implement accessors getter, setter and caller. (version 0.1.x) (Some V8Variants were advanced to 0.1.x .)


# API

See the [API documentation](https://github.com/idobatter/node-win32ole/wiki) in the wiki.


# BUILDING

This project uses VC++ 2008 Express (or later) and Python 2.6 (or later) .
(When using Python 2.5, it needs [multiprocessing 2.5 back port](http://pypi.python.org/pypi/multiprocessing/) .) It needs neither ATL nor MFC.

Bulding also requires node-gyp to be installed. You can do this with npm:

    npm install -g node-gyp

To obtain and build the bindings:

    git clone git://github.com/idobatter/node-win32ole.git
    cd node-win32ole
    node-gyp configure
    node-gyp build

You can also use [`npm`](https://github.com/isaacs/npm) to download and install them:

    npm install win32ole


# TESTS

[mocha](https://github.com/visionmedia/mocha) is required to run unit tests.

    npm install -g mocha
    nmake /a test


# CONTRIBUTORS

* [idobatter](https://github.com/idobatter)


# ACKNOWLEDGEMENTS

Inspired [Win32OLE](http://www.ruby-doc.org/stdlib/libdoc/win32ole/rdoc/)


# LICENSE

`node-win32ole` is [BSD licensed](https://github.com/idobatter/node-win32ole/raw/master/LICENSE).
