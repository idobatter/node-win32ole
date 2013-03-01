var win32ole = require('win32ole');
win32ole.print('activex_filesystemobject_sample\n');

var testfile = 'examples\\activex_filesystemobject_sample.js';

var activex_filesystemobject_sample = function(){
  var withReadFile = function(filename, callback){
    var fso = new ActiveXObject('Scripting.FileSystemObject');
    var fullpath = fso.GetAbsolutePathName(filename);
    var file = fso.OpenTextFile(fullpath, 1, false); // open to read
    try{
      callback(file);
    }finally{
      file.Close();
    }
  };
  var withEachLine = function(filename, callback){
    withReadFile(filename, function(file){
/*
 In ParseUnaryExpression() < v8/src/parser.cc >
 v8::Object::ToBoolean() is called directly for unary operator '!'
 instead of v8::Object::valueOf()
 so NamedPropertyHandler will not be called
 Local<Boolean> ToBoolean(); // How to fake ? override v8::Value::ToBoolean
*/
//    while(file.AtEndOfStream != true) // It works. (without unary operator !)
//    while(!file.AtEndOfStream) // It does not work.
      while(!file.AtEndOfStream._) // *** It works. oops!
        callback(file.ReadLine());
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
