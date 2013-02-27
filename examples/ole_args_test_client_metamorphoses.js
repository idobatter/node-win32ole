/*
How to use it on Windows 7/Vista:
1. register server
  1-1. Enable your PC's Administrator account.
  1-2. runas /user:administrator "python ole_args_tester_server.py"
2. node ole_args_test_client_metamorphoses.js
  ( win32ole.client.Dispatch must be overrided by metamorphoses proxy )
*/

var win32ole = require('win32ole');
win32ole.print('ole_args_test_client_metamorphoses\n');

var ole_args_test_client_metamorphoses = function(){
  console.log('test');
  var cl = win32ole.client.Dispatch('OLEArgsTester.Server');
  console.log('connected (and metamorphoses proxy when wrapped)');
  console.log('test caller');
  cl.Init(); // cl.call('Init');
  console.log('test arguments (call)');
  console.log(cl.call('Test', ['a', 'b', 'c', 'd']));
  console.log('test arguments (__noSuchMethod__)');
  console.log(cl.Test('a1'));
  console.log(cl.Test('a2', 'b2'));
  console.log(cl.Test('a3', 'b3', 'c3'));
  console.log(cl.Test('a4', 'b4', 'c4', 'd4'));
  console.log('test getter (call)');
  console.log(cl.call('GetSubName'));
  console.log('test arguments (__noSuchMethod__) * this is not a setter *');
  console.log(cl.SetSubName('cba'));
  console.log('test getter (*** __noSuchMethod__ ***)');
  console.log(cl.GetSubName());
//  console.log('test getter (*** __noSuchGetter__ ***)');
//  console.log(cl.GetSubName._); // ERROR: method MUST NOT be called as getter
  console.log('test setter attr (__noSuchSetter__)');
  console.log(cl.subname = 'zyx');
  console.log('test getter attr (*** __noSuchGetter__ ***)');
  console.log(cl.subname); // [object V8Variant] (called cl.subname.inspect)
  console.log(cl.subname.toUtf8()); // instance method (obsoleted)
  // discussed about __noSuchProperty__
  // https://mail.mozilla.org/pipermail/es-discuss/2010-October/011930.html
  console.log(cl.subname._); // (obsoleted)
  console.log('test getter accessor (*** __noSuchMethod__ ***)');
  console.log(cl.subname()); // getter may be called as method (obsoleted)
  console.log('test getter (*** __noSuchMethod__ ***)');
  console.log(cl.GetSubName());
  console.log('quit');
  cl.Quit(); // cl.call('Quit');
  console.log('completed');
};

try{
  ole_args_test_client_metamorphoses();
}catch(e){
  console.log('*** exception cached ***\n' + e);
}
