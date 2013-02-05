var win32ole = require('win32ole');
win32ole.print('wsh_sample\n');
console.log(win32ole.version());
var path = require('path');
var cwd = path.join(win32ole.MODULEDIRNAME, '..');

var fs = require('fs');
var tmpdir = path.join(cwd, 'test/tmp');
if(!fs.existsSync(tmpdir)) fs.mkdirSync(tmpdir);
var outfile = path.join(tmpdir, 'wsh_sample.txt');

var wsh_sample = function(filename){
  var sh = win32ole.client.Dispatch('WScript.Shell', 'C');
  console.log('sh:');
  console.log(require('util').inspect(sh, true, null, true));
  try{
    console.log('notepad');
    sh.call('run', ['notepad.exe']); // ['notepad.exe', 1, true]
    console.log('ok');
  }catch(e){
    console.log('(exception catched)' + e);
  }
  sh.call('run', ['ping.exe -t 127.0.0.1']);
  console.log('completed');
  sh = null;
};

win32ole.client = new win32ole.Client;
try{
  wsh_sample(outfile);
}catch(e){
  console.log('*** exception cached ***\n' + e);
}
win32ole.client.Finalize(); // must be called (version 0.0.x)
