var win32ole = require('win32ole');
win32ole.print('uncfinder_sample');
console.log(win32ole.version());
var path = require('path');
var cwd = path.join(win32ole.MODULEDIRNAME, '..');

var fs = require('fs');
var tmpdir = path.join(cwd, 'test/tmp');
if(!fs.existsSync(tmpdir)) fs.mkdirSync(tmpdir);
var outfile = path.join(tmpdir, 'uncfinder_sample.txt');

var uncfinder_sample = function(filename){
  var wnt = win32ole.client.Dispatch('WinNTSystemInfo', 'C');
  console.log('ComputerName:');
  console.log(wnt.get('ComputerName').toUtf8());
  console.log('DomainName:');
  console.log(wnt.get('DomainName').toUtf8());
  console.log('PDC:');
  console.log(wnt.get('PDC').toUtf8());
  console.log('UserName:');
  console.log(wnt.get('UserName').toUtf8());
};

win32ole.client = new win32ole.Client;
try{
  uncfinder_sample(outfile);
}catch(e){
  console.log('*** exception cached ***\n' + e);
}
win32ole.client.Finalize(); // must be called (version 0.0.x)
