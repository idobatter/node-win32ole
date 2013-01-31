// Inspired by https://github.com/TooTallNate/node-sqlite3
var win32ole = module.exports = exports =
  require('../build/Release/node_win32ole.node');
win32ole.get_package_version = function(){
  var fs = require('fs');
  try{
    win32ole.VERSION = JSON.parse(
      fs.readFileSync('./package.json', 'utf8'))['version'];
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
var Statement = win32ole.Statement;
inherits(Statement, EventEmitter);

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
    trace.extendTrace(Statement.prototype, 'all');
    trace.extendTrace(Statement.prototype, 'map');
    isVerbose = true;
  }
  return this;
};
