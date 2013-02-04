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
var Statement = win32ole.Statement;

inherits(V8Variant, EventEmitter);
inherits(Statement, EventEmitter);

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

Statement.prototype.map = function(){
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
    trace.extendTrace(V8Variant.prototype, 'get');
    trace.extendTrace(V8Variant.prototype, 'set');
    trace.extendTrace(V8Variant.prototype, 'call');
    trace.extendTrace(V8Variant.prototype, 'map');
    trace.extendTrace(Statement.prototype, 'Dispatch');
    trace.extendTrace(Statement.prototype, 'map');
    isVerbose = true;
  }
  return this;
};
