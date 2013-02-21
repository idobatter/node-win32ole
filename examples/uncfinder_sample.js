var win32ole = require('win32ole');
win32ole.print('uncfinder_sample\n');
console.log(win32ole.version());
var path = require('path');
var cwd = path.join(win32ole.MODULEDIRNAME, '..');

var fs = require('fs');
var tmpdir = path.join(cwd, 'test/tmp');
if(!fs.existsSync(tmpdir)) fs.mkdirSync(tmpdir);
var outfile = path.join(tmpdir, 'uncfinder_sample.txt');

var uncfinder_sample = function(filename){
  var wnt = win32ole.client.Dispatch('WinNTSystemInfo');
  console.log('ComputerName:');
  console.log(wnt.ComputerName._); // ***
  console.log('DomainName:');
  console.log(wnt.DomainName._); // ***
  console.log('PDC:');
  console.log(wnt.PDC._); // ***
  console.log('UserName:');
  console.log(wnt.UserName._); // ***
};

try{
  uncfinder_sample(outfile);
}catch(e){
  console.log('*** exception cached ***\n' + e);
}
