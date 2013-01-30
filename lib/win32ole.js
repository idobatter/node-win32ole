var win32ole = module.exports = exports =
  require('../build/Release/node_win32ole.node');
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
