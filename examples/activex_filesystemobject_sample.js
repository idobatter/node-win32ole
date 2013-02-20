var win32ole = require('win32ole');
var util = require('util');
win32ole.print('activex_filesystemobject_sample\n');

var testfile = 'examples\\activex_filesystemobject_sample.js';

var activex_filesystemobject_sample = function(){
  var withReadFile = function(filename, callback){
//    var fullpath = fso.GetAbsolutePathName(filename).toUtf8();
    var fullpath = fso.call('GetAbsolutePathName', [filename]).toUtf8();
    console.log(fullpath);
//    var file = fso.OpenTextFile(fullpath, 1, false); // open to read
    var file = fso.call('OpenTextFile', [fullpath, 1, false]); // open to read
    try{
      callback(file);
    }finally{
//      file.Close();
      file.call('Close');
    }
  };
  var withEachLine = function(filename, callback){
    withReadFile(filename, function(file){
//      while(!file.AtEndOfStream)
//        callback(file.ReadLine());
      while(!file.get('AtEndOfStream').toBoolean())
        callback(file.call('ReadLine').toUtf8());
    });
  };
  var fso = new ActiveXObject('Scripting.FileSystemObject');
  withEachLine(testfile, function(line){
    console.log(line);
  });
};

try{
  activex_filesystemobject_sample();
}catch(e){
  console.log('*** exception cached ***\n' + e);
}
