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
  console.log('test0');
  var cl = win32ole.client.Dispatch('OLEArgsTester.Server');
  console.log('test');
  console.log(cl.call('Test', ['a', 'b', 'c', 'd']).toUtf8());
  console.log(cl.Test('a1').toUtf8());
  console.log(cl.Test('a2', 'b2').toUtf8());
  console.log(cl.Test('a3', 'b3', 'c3').toUtf8());
  console.log(cl.Test('a4', 'b4', 'c4', 'd4').toUtf8());
  console.log('quit');
  cl.Quit(); // cl.call('Quit');
  console.log('completed');
};

try{
  ole_args_test_client_metamorphoses();
}catch(e){
  console.log('*** exception cached ***\n' + e);
}
