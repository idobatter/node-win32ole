/*
  WMI Scripting Library Object Model
  http://technet.microsoft.com/en-us/library/ee198924.aspx
  Scripting API Objects
  http://msdn.microsoft.com/en-us/library/aa393259%28v=vs.85%29.aspx
  SWbemObjectSet object (Windows)
  http://msdn.microsoft.com/en-us/library/aa393762%28v=vs.85%29.aspx
  SWbemObjectSet::ItemIndex method (Windows)
  http://msdn.microsoft.com/en-us/library/aa826600%28v=vs.85%29.aspx
  Enumerating WMI
  http://msdn.microsoft.com/en-us/library/aa390386%28v=vs.85%29.aspx
  SWbemObjectSet is received from SWbemServices.ExecQuery
  COM API for WMI
  http://msdn.microsoft.com/en-us/library/aa389276%28v=vs.85%29.aspx
  (C++) IEnumWbemClassObject is received from IWbemServices::ExecQuery
*/
var win32ole = require('win32ole');
win32ole.print('wmi_sample\n');
console.log(win32ole.version());
var path = require('path');
var cwd = path.join(win32ole.MODULEDIRNAME, '..');

var fs = require('fs');
var tmpdir = path.join(cwd, 'test/tmp');
if(!fs.existsSync(tmpdir)) fs.mkdirSync(tmpdir);
var outfile = path.join(tmpdir, 'wmi_sample.txt');

var wmi_sample = function(filename){
  var locator = win32ole.client.Dispatch('WbemScripting.SWbemLocator');
  var svc = locator.call('ConnectServer', ['.', 'root/cimv2']); // . = self PC
  try{
    if(fs.existsSync(filename)) fs.unlinkSync(filename);
    var query = "select * from Win32_Process where Name like '%explore%'";
    query += " or Name='rundll32.exe' or Name='winlogon.exe'";
    var procset = svc.call('ExecQuery', [query]);
    console.log('procset is a ' + procset.isA()); // VT_DISPATCH=SWbemObjectSet
    var count = procset.get('Count').toInt32();
    console.log('count = ' + count);
    console.log(' ImageName, ProcessId, VirtualSize, Threads, Description,');
    console.log(' [CommandLine],');
    console.log(' [ImagePath]');
    for(var i = 0; i < count; ++i){
      var proc = procset.call('ItemIndex', [i]);
      var safeutf8 = function(obj, propertyname){
        var property = obj.get(propertyname);
        return property.isA() == 1 ? 'NULL' : property.toUtf8();
      };
      win32ole.printACP(
        '-> ' + safeutf8(proc, 'Name')
        + ', ' + proc.get('ProcessId').toInt32()
        + ', ' + safeutf8(proc, 'VirtualSize')
        + ', ' + proc.get('ThreadCount').toInt32()
        + ', ' + safeutf8(proc, 'Description')
        + '\n   [' + safeutf8(proc, 'CommandLine')
        + ']\n   [' + safeutf8(proc, 'ExecutablePath')
        + ']\n');
    }
  }catch(e){
    console.log('(exception catched)\n' + e);
  }

  console.log('completed');
};

try{
  wmi_sample(outfile);
}catch(e){
  console.log('*** exception cached ***\n' + e);
}
