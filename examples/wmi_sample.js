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

var safeutf8 = function(obj, propertyname){
  var property = obj.get(propertyname); // ***
  // return property.isA() == 1 ? 'NULL' : property;
  // return property.vtName() == 'VT_NULL' ? 'NULL' : property;
  return property == null ? 'NULL' : property;
};

var get_value_from_key = function(kv, key){
  // search from kv where Item['Name'] == key, return found Item['Value']
  try{
    var item = kv.Item(key);
    try{
      return safeutf8(item, 'Value');
    }catch(e){
      return key + ': *** BUG type missmatch (process int32 or array) ***';
    }
  }catch(e){
    return key + ': *** BUG no object (process OCVariant::AutoWrap) ***';
  }
};

var wmi_sample = function(filename){
  var locator = win32ole.client.Dispatch('WbemScripting.SWbemLocator');
  var svr = locator.ConnectServer('.', 'root/cimv2'); // . = self PC
  try{
    if(fs.existsSync(filename)) fs.unlinkSync(filename);
    console.log('*** get processes (partial) ***');
    var query = "select * from Win32_Process where Name like '%explore%'";
    query += " or Name='rundll32.exe' or Name='winlogon.exe'";
    var procset = svr.ExecQuery(query);
    console.log('procset is a ' + procset.vtName()); // ( SWbemObjectSet )
    var count = procset.Count;
    console.log('count = ' + count);
    console.log(' ImageName, ProcessId, VirtualSize, Threads, Description,');
    console.log(' [ImagePath]');
    console.log(' [CommandLine],');
    for(var i = 0; i < count; ++i){
      var proc = procset.ItemIndex(i);
      var imgpath = safeutf8(proc, 'ExecutablePath');
      var cmdline = safeutf8(proc, 'CommandLine');
      var size = imgpath.length + 1;
      if(imgpath.match(/[\s]+/ig)) size += 2;
      win32ole.printACP(
        '-> ' + safeutf8(proc, 'Name')
        + ', ' + proc.ProcessId
        + ', ' + safeutf8(proc, 'VirtualSize')
        + ', ' + proc.ThreadCount
        + ', ' + safeutf8(proc, 'Description')
        + '\n   [' + imgpath
        + ']\n   [' + cmdline.substring(size)
        + ']\n');
    }
    console.log('*** get services (first 10 items) ***');
    var svcset = svr.ExecQuery('select * from Win32_Service');
    count = svcset.Count;
    console.log('count = ' + count);
    for(var i = 0; i < 10; ++i){
      var svc = svcset.ItemIndex(i);
      var q = svc.Qualifiers_._; // ***
      win32ole.printACP(
        '-> ' + get_value_from_key(q, 'provider')
        + '\n   [' + get_value_from_key(q, 'UUID')
        + ']\n');
      var p = svc.Properties_._; // ***
      win32ole.printACP(
        '   ' + get_value_from_key(p, 'Name')
        + '\n   [' + get_value_from_key(p, 'PathName')
        + ']\n');
      var np = function(m, na){
        for(var j = 0; j < na.length; ++j){
          console.log('    method Qualifiers_: ' + na[j]);
          var me = m.Item(na[j]);
          var mq = me.Qualifiers_._; // ***
          win32ole.printACP(
            '     [' + get_value_from_key(mq, 'Override')
            + ']\n     [' + get_value_from_key(mq, 'Static') // Boolean
            + ']\n     [' + get_value_from_key(mq, 'MappingStrings') // Array
            + ']\n     [' + get_value_from_key(mq, 'ValueMap') // Array
            + ']\n');
        }
      };
      var m = svc.Methods_;
      console.log('   methods: ' + m.Count);
      if(true){
        // do nothing here because there are too many bugs (get_value_from_key)
        // np(m, [me.Name for me in svc.Methods_]); // dummy code (as python)
      }else{
        // for each me.Name is not taken from svc.Methods_ ... (version 0.0.x)
        np(m, ['StartService', 'StopService', 'PauseService', 'ResumeService',
          'InterrogateService', 'UserControlService',
          'Create', 'Change', 'ChangeStartMode', 'Delete',
          'GetSecurityDescriptor', 'SetSecurityDescriptor']);
      }
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
