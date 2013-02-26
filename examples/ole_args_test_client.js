/*
How to use it on Windows 7/Vista:
1. register server
  1-1. Enable your PC's Administrator account.
  1-2. runas /user:administrator "python ole_args_tester_server.py"
2. node ole_args_test_client.js
*/

var win32ole = require('win32ole');
win32ole.print('ole_args_test_client\n');

var ole_args_test_client = function(){
  console.log('create connection');
  var cl = win32ole.client.Dispatch('OLEArgsTester.Server');
  console.log('connected');
  cl.call('Init');
  console.log('do A');
  console.log(cl.subname = 'YoY');
  console.log('do B');
  console.log(cl.subname);
  console.log('do 1');
  console.log(cl.call('Test', ['a']));
  console.log('do 2');
  console.log(cl.call('Test', ['a1', 'b2']));
  console.log('do 3');
  console.log(cl.call('Test', ['aa1', 'bb2', 'cc3']));
  console.log('do 4');
  console.log(cl.call('Test', ['aaa1', 'bbb2', 'ccc3', 'ddd4']));
  console.log('quit');
  cl.call('Quit');
  console.log('disconnected');

  // process.on('exit'...) will *NOT* be called when there is not enough memory
  var force_gc = function(){
    // cl = null; // V8Variant.Dispose() will be called when cl is *NOT* null
    // win32ole.client = null; // Client.Dispose() will be called only if null
    console.log('a');
    for(var a = [], i = 0; i < 869461; i++){ a[i] = 'dummydata'; } // GC test
    console.log('b');
    for(var b = [], j = 0; j < 2944; j++){ b[j] = 'bbbbbbbb'; } // GC test
    console.log('c');
    for(var c = [], k = 0; k < 848; k++){ c[k] = 'cccccccc'; } // GC test
    console.log('d');
    for(var d = [], l = 0; l < 358; l++){ d[l] = 'dddddddd'; } // GC test
  }

  // force GC (disposer will *NOT* be called if GC is not run)
  var forceGCmode = 0; // 0, 1, 2, 3
  switch(forceGCmode){
  case 1: force_gc(); break;
  case 2: win32ole.force_gc_extension(1); break;
  case 3: win32ole.force_gc_internal(1); break;
  default: break; // do nothing
  }

  console.log('completed');
};

try{
  ole_args_test_client();
}catch(e){
  console.log('*** exception cached ***\n' + e);
}
