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
  var sh = win32ole.client.Dispatch('WScript.Shell');
  console.log('sh:');
  console.log(require('util').inspect(sh, true, null, true));

  try{

    if(fs.existsSync(filename)) fs.unlinkSync(filename);
    win32ole.print('notepad ...');
    // arg1=1: movetop (default), 2: minimize, 3: maximize, 4: no movetop
    //      5: movetop, 6: minimize, 7: minimize
    // arg2=false: Async (default), true: Sync
    sh.call('run', ['notepad.exe', 5, false]); // must be Async (for SendKeys)
    win32ole.sleep(3000, true, false);
    sh.call('SendKeys', ['Congratulations!{ENTER}']); // 'run' option must be 5
    win32ole.sleep(200);
    sh.call('SendKeys', ["*** DON'T TOUCH THIS WINDOW ***{ENTER}"]);
    win32ole.sleep(3000);
    sh.call('SendKeys', ["Saving filename will be entered automatically."]);
    win32ole.sleep(3000);
    sh.call('SendKeys', ['%f']); // ALT-F (File)
    win32ole.sleep(500);
    sh.call('SendKeys', ['a']); // SaveAs
    win32ole.sleep(5000);
    sh.call('SendKeys', [filename + '{ENTER}']);
    win32ole.sleep(1000);
    sh.call('SendKeys', ['%{F4}']); // ALT-F4 (Exit)
    win32ole.sleep(100);
    win32ole.print('ok\n');

    win32ole.print('ping ...');
    sh.call('run', ['ping.exe 127.0.0.1', 4, true]); // wait for close (Sync)
    win32ole.print('ok\n');

  }catch(e){
    console.log('(exception catched)' + e);
  }

  var shellexec = function(sh, cmd, callback){
    console.log('sh.Exec(' + cmd + ')');
    var stat = sh.call('Exec', [cmd]);
//    while(stat.get('Status').toInt32() == 0) win32ole.sleep(100, true, true);
    var so = stat.get('StdOut');
    while(!so.get('AtEndOfStream').toBoolean())
      callback(so.call('ReadLine').toUtf8());
    console.log('code = ' + stat.get('ExitCode').toInt32());
  }

  var cmd = 'reg query "HKLM\\Software\\Microsoft\\Internet Explorer"';
  shellexec(sh, cmd, function(line){
    if(line.match(/([^\s]*version[^\s]*)[\s]+([^\s]+)[\s]+([^\s]+)/ig))
      console.log(RegExp.$1 + ',' + RegExp.$2 + ',' + RegExp.$3);
  });

  shellexec(sh, 'ipconfig.exe', function(line){ console.log(line); });

  win32ole.print('writing to eventlog ...');
  var name = 'node-win32ole (' + win32ole.VERSION + ') ';
  // the first argument value
  //  0: EVENTLOG_SUCCESS
  //  1: EVENTLOG_ERROR_TYPE
  //  2: EVENTLOG_WARNING_TYPE
  //  4: EVENTLOG_INFORMATION_TYPE
  //  8: EVENTLOG_AUDIT_SUCCESS
  // 16: EVENTLOG_AUDIT_FAILURE
  // the 3rd argument '.' means 'self computer name' to send LogEvent target
  sh.call('LogEvent', [4, name + 'installed', '.']);
  sh.call('LogEvent', [2, name + 'warning test', '.']);
  sh.call('LogEvent', [1, name + 'error test', '.']);
  sh.call('LogEvent', [0, name + 'success', '.']);
  win32ole.print('ok\n');

  console.log('completed');
};

try{
  wsh_sample(outfile);
}catch(e){
  console.log('*** exception cached ***\n' + e);
}
