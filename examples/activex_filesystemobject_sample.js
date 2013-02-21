var win32ole = require('win32ole');
win32ole.print('activex_filesystemobject_sample\n');

var testfile = 'examples\\activex_filesystemobject_sample.js';

var activex_filesystemobject_sample = function(){
  var withReadFile = function(filename, callback){
    var fso = new ActiveXObject('Scripting.FileSystemObject');
    var fullpath = fso.GetAbsolutePathName(filename); // .toUtf8(); // ***
    var file = fso.OpenTextFile(fullpath, 1, false); // open to read
    try{
      callback(file);
    }finally{
      file.Close();
    }
  };
  var withEachLine = function(filename, callback){
    withReadFile(filename, function(file){
      while(!file.AtEndOfStream._.toBoolean()) // ***
        callback(file.ReadLine().toUtf8()); // ***
    });
  };
  withEachLine(testfile, function(line){
    console.log(line);
  });
};

try{
  activex_filesystemobject_sample();
}catch(e){
  console.log('*** exception cached ***\n' + e);
}
