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
  var sh = win32ole.client.Dispatch('WScript.Shell', '.ACP'); // locale
  console.log('sh:');
  console.log(require('util').inspect(sh, true, null, true));

  try{

    if(fs.existsSync(filename)) fs.unlinkSync(filename);
    win32ole.print('notepad ...');
    // arg1=1: movetop (default), 2: minimize, 3: maximize, 4: no movetop
    //      5: movetop, 6: minimize, 7: minimize
    // arg2=false: Async (default), true: Sync
    sh.call('run', ['notepad.exe', 5]); // , true *** bug: force exit ***
    win32ole.print('waiting 2 seconds ...');
    sleep(2000);
    sh.call('SendKeys', ['Congratulations!']); // must call 'run' option 5
    sleep(200);
    sh.call('SendKeys', ['%f']); // ALT-F (File)
    sleep(500);
    sh.call('SendKeys', ['a']); // SaveAs
    sleep(3000);
    sh.call('SendKeys', [filename + '{ENTER}']);
    sleep(1000);
    sh.call('SendKeys', ['%{F4}']); // ALT-F4 (Exit)
    sleep(100);
    win32ole.print('ok\n');

    win32ole.print('ping ...');
    sh.call('run', ['ping.exe -t 127.0.0.1', 4]);
    win32ole.print('ok\n');

  }catch(e){
    console.log('(exception catched)' + e);
  }

  var shellexec = function(sh, cmd, callback){
    console.log('sh.Exec(' + cmd + ')');
    var stat = sh.call('Exec', [cmd]);
    var so = stat.get('StdOut');
    while(!so.get('AtEndOfStream').toBoolean())
      callback(so.call('ReadLine').toUtf8());
/*
    while(stat.get('Status').toInt32() == 0){
      console.log('waiting ...');
      sleep(100);
    }
*/
    console.log('code = ' + stat.get('ExitCode').toInt32());
  }

  var cmd = 'reg query "HKLM\\Software\\Microsoft\\Internet Explorer"';
  shellexec(sh, cmd, function(line){
    if(line.match(/([^\s]*version[^\s]*)[\s]+([^\s]+)[\s]+([^\s]+)/ig))
      console.log(RegExp.$1 + ',' + RegExp.$2 + ',' + RegExp.$3);
  });

  shellexec(sh, 'ipconfig.exe', function(line){ console.log(line); });

  win32ole.print('writing to eventlog ...');
  // 0: EVENTLOG_SUCCESS
  // 1: EVENTLOG_ERROR_TYPE
  // 2: EVENTLOG_WARNING_TYPE
  // 4: EVENTLOG_INFORMATION_TYPE
  // 8: EVENTLOG_AUDIT_SUCCESS
  // 16: EVENTLOG_AUDIT_FAILURE
  sh.call('LogEvent', [4, 'node-win32ole installed']);
  sh.call('LogEvent', [2, 'node-win32ole warning test']);
  sh.call('LogEvent', [1, 'node-win32ole error test']);
  sh.call('LogEvent', [0, 'node-win32ole success']);
  win32ole.print('ok\n');

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
