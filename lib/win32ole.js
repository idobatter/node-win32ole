// Inspired by https://github.com/TooTallNate/node-sqlite3
var win32ole = module.exports = exports =
  require('../build/Release/node_win32ole.node');
win32ole.MODULEDIRNAME = __dirname;
win32ole.get_package_version = function(){
  var fs = require('fs');
  var path = require('path');
  var fn = path.join(win32ole.MODULEDIRNAME, '../package.json');
  try{
    win32ole.VERSION = JSON.parse(fs.readFileSync(fn, 'utf8'))['version'];
  }catch(e){
    win32ole.VERSION = e;
  }
};
win32ole.get_package_version();

/*
  exchange (VT type <-> v8 type)
  VT_ERROR -> *** ( Int32 ) ***
  VT_EMPTY -> *** ( Empty or Undefined ) ***
  VT_SAFEARRAY <-> VT_SAFEARRAY, (VT_BYREF | VT_SAFEARRAY)
  VT_VARIANT <-> VT_VARIANT, VT_USERDEFINED, VT_ARRAY, (VT_BYREF | VT_VARIANT)
  exchange (VT type <-> C/C++ type <-> v8 type)
  VT_DISPATCH, VT_UNKNOWN <-> OCVariant <-> V8Variant
  VT_BSTR (VT_CLSID, VT_FILETIME, VT_LPWSTR, VT_LPSTR,
    VT_DECIMAL, VT_ERROR, VT_BSTR) <-> string <-> Utf8Value
  VT_DATE <-> double <-> Date
  VT_I2 -> int16_t (short) -> Int32
  VT_UI2 -> uint16_t (ushort) -> Int32
  VT_I4 (VT_INT, VT_I4) <-> int32_t (long) <-> Int32
  VT_UI4 (VT_UINT, VT_UI4) -> uint32_t (ulong) -> Int32
  VT_I8 -> int64_t (long long) -> Number
  VT_UI8 -> uint64_t (ulonglong) -> Number
  VT_R4 -> float -> Number
  VT_R8 (VT_CY, VT_R8) <-> double <-> Number
  VT_I1 (VT_UI1, VT_I1) -> char (uchar) -> *** ( Int32 ) ***
  VT_BOOL <-> bool <-> Boolean
  VT_NULL <-> NULL <-> Null
*/

win32ole.vt_enum = {
  VT_EMPTY: 0, VT_NULL: 1,
  VT_I2: 2, VT_I4: 3, VT_R4: 4, VT_R8: 5, VT_CY: 6, VT_DATE: 7, VT_BSTR: 8,
  VT_DISPATCH: 9, VT_ERROR: 10, VT_BOOL: 11, VT_VARIANT: 12, VT_UNKNOWN: 13,
  VT_DECIMAL: 14, // 15
  VT_I1: 16, VT_UI1: 17, VT_UI2: 18, VT_UI4: 19,
  VT_I8: 20, VT_UI8: 21, VT_INT: 22, VT_UINT: 23,
  VT_VOID: 24, VT_HRESULT: 25, VT_PTR: 26, VT_SAFEARRAY: 27, VT_CARRAY: 28,
  VT_USERDEFINED: 29, VT_LPSTR: 30, VT_LPWSTR: 31, // 32-35
  VT_RECORD: 36, // 37-63
  VT_FILETIME: 64, VT_BLOB: 65, VT_STREAM: 66, VT_STORAGE: 67,
  VT_STREAM_OBJECT: 68, VT_STORED_OBJECT: 69, VT_BLOB_OBJECT: 70,
  VT_CF: 71, VT_CLSID: 72, // 73-4094
  VT_BSTR_BLOB: 4095, // 0x0fff
  VT_VECTOR: 4096, // flag 0x1000
  VT_ARRAY: 8192, // flag 0x2000
  VT_BYREF: 16384, // flag 0x4000
  VT_RESERVED: 32768, // flag 0x8000
  VT_ILLEGAL: 65535, // -1 0xffff
  VT_ILLEGALMASKED: 4095, // 0x0fff *** caution ***
  VT_TYPEMASK: 4095 // 0x0fff *** caution ***
};

win32ole.vt_names = (function(){
  var names = {};
  var vte = win32ole.vt_enum;
  var xkey = 'VT_BSTR_BLOB';
  var xnum = vte[xkey];
  for(var k in vte) if(vte[k] != xnum) names[vte[k]] = k; // must use typemask
  names[xnum] = xkey;
  return names;
})();

var util = require('util');
var EventEmitter = require('events').EventEmitter;

function errorCallback(args){
  if(typeof args[args.length - 1] === 'function'){
    var callback = args[args.length - 1];
    return function(err){ if(err) callback(err); }
  }
}

function inherits(target, source){
  for(var k in source.prototype) target.prototype[k] = source.prototype[k];
}

var isVerbose = false;
var supportedEvents = ['trace', 'profile'];
var V8Variant = win32ole.V8Variant;
var Client = win32ole.Client;

inherits(V8Variant, EventEmitter);
inherits(Client, EventEmitter);

V8Variant.prototype.map = function(){
  var params = Array.prototype.slice.call(arguments);
  var callback = params.pop();
  params.push(function(err, rows){
    if(err) return callback(err);
    var result = {};
    if(rows.length){
      var keys = Object.keys(rows[0]), key = keys[0];
      if(keys.length > 2){ // Value is an Object
        for(var i = 0; i < rows.length; i++)
          result[rows[i][key]] = rows[i];
      }else{ // Value is a plain value
        var value = keys[i];
        for(var i = 0; i < rows.length; i++)
          result[rows[i][key]] = rows[i][value];
      }
    }
    callback(err, result);
  });
  return this.all.apply(this, params);
};

Client.prototype.map = function(){
  var params = Array.prototype.slice.call(arguments);
  var callback = params.pop();
  params.push(function(err, rows){
    if(err) return callback(err);
    var result = {};
    if(rows.length){
      var keys = Object.keys(rows[0]), key = keys[0];
      if(keys.length > 2){ // Value is an Object
        for(var i = 0; i < rows.length; i++)
          result[rows[i][key]] = rows[i];
      }else{ // Value is a plain value
        var value = keys[i];
        for(var i = 0; i < rows.length; i++)
          result[rows[i][key]] = rows[i][value];
      }
    }
    callback(err, result);
  });
  return this.all.apply(this, params);
};

// Save the stack trace over EIO callbacks.
win32ole.verbose = function(){
  if(!isVerbose){
    var trace = require('./trace');
    trace.extendTrace(V8Variant.prototype, 'isA');
    trace.extendTrace(V8Variant.prototype, 'vtName');
    trace.extendTrace(V8Variant.prototype, 'toBoolean'); // *** p.
    trace.extendTrace(V8Variant.prototype, 'toInt32'); // *** p.
    trace.extendTrace(V8Variant.prototype, 'toInt64'); // *** p.
    trace.extendTrace(V8Variant.prototype, 'toNumber'); // *** p.
    trace.extendTrace(V8Variant.prototype, 'toDate'); // *** p.
    trace.extendTrace(V8Variant.prototype, 'toUtf8'); // *** p.
    trace.extendTrace(V8Variant.prototype, 'toValue');
    trace.extendTrace(V8Variant.prototype, 'call');
    trace.extendTrace(V8Variant.prototype, 'get');
    trace.extendTrace(V8Variant.prototype, 'set');
    trace.extendTrace(V8Variant.prototype, 'map');
    trace.extendTrace(Client.prototype, 'Dispatch');
    trace.extendTrace(Client.prototype, 'map');
    isVerbose = true;
  }
  return this;
};

win32ole.client = new win32ole.Client;
process.on('exit', function(){
  win32ole.client.Finalize();
  // win32ole.print('EXIT\n');
});

global.ActiveXObject = function(args){
  return win32ole.client.Dispatch(args);
};
