var win32ole = require('win32ole');
win32ole.print('wsh_sample\n');
console.log(win32ole.version());
var path = require('path');
var cwd = path.join(win32ole.MODULEDIRNAME, '..');

var fs = require('fs');
var tmpdir = path.join(cwd, 'test/tmp');
if(!fs.existsSync(tmpdir)) fs.mkdirSync(tmpdir);
var outfile = path.join(tmpdir, 'wsh_sample.txt');

var sleep = function(milliSeconds){
  var startTime = new Date().getTime();
  while(new Date().getTime() < startTime + milliSeconds);
}

var wsh_sample = function(filename){
  var sh = win32ole.client.Dispatch('WScript.Shell', 'C');
  console.log('sh:');
  console.log(require('util').inspect(sh, true, null, true));
  try{
    win32ole.print('notepad ...');
    // arg1=1: movetop (default), 2: minimize, 3: maximize, 4: no movetop
    //      5: movetop, 6: minimize, 7: minimize
    // arg2=false: Async (default), true: Sync
    sh.call('run', ['notepad.exe', 4]);
    win32ole.print('ok\n');
    win32ole.print('ping ...');
    sh.call('run', ['ping.exe -t 127.0.0.1', 4]);
    win32ole.print('ok\n');
  }catch(e){
    console.log('(exception catched)' + e);
  }
  var cmd = 'reg query "HKLM\\Software\\Microsoft\\Internet Explorer"';
  console.log('sh.Exec(' + cmd + ')');
  var stat = sh.call('Exec', [cmd]);
/*
  while(stat.get('Status').toInt32() == 0){
    console.log('waiting ...');
    sleep(100);
  }
*/
  while(!stat.get('StdOut').get('AtEndOfStream').toInt32()){ // toBoolean
    var line = stat.get('StdOut').call('ReadLine').toUtf8();
    if(line.match(/([^\s]*version[^\s]*)[\s]+([^\s]+)[\s]+([^\s]+)/ig))
      console.log(RegExp.$1 + ',' + RegExp.$2 + ',' + RegExp.$3);
  }
  console.log('code = ' + stat.get('ExitCode').toInt32());
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
