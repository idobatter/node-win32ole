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
  console.log('do 1');
  console.log(cl.call('Test', ['a']).toUtf8());
  console.log('do 2');
  console.log(cl.call('Test', ['a1', 'b2']).toUtf8());
  console.log('do 3');
  console.log(cl.call('Test', ['aa1', 'bb2', 'cc3']).toUtf8());
  console.log('do 4');
  console.log(cl.call('Test', ['aaa1', 'bbb2', 'ccc3', 'ddd4']).toUtf8());
  console.log('connection');
  cl.call('Quit');
  console.log('disconnected');
};

try{
  ole_args_test_client();
}catch(e){
  console.log('*** exception cached ***\n' + e);
}
win32ole.client.Finalize(); // must be called (version 0.0.x)
