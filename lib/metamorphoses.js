var Proxy = require('node-proxy');

var PlaneProxyHandler = function(){}; // abstract (prototype of prototype)

PlaneProxyHandler.prototype.obj = null; // Function or Object (overrided later)

// fundamental traps

PlaneProxyHandler.prototype.getOwnPropertyDescriptor = function(name){
  var desc = Object.getOwnPropertyDescriptor(this.obj, name);
  // a trapping proxy's properties must always be configurable
  if(desc !== undefined){ desc.configurable = true; }
  return desc;
};

PlaneProxyHandler.prototype.getPropertyDescriptor = function(name){
  var desc = Object.getPropertyDescriptor(this.obj, name); // !ES5
  // a trapping proxy's properties must always be configurable
  if(desc !== undefined){ desc.configurable = true; }
  return desc;
};

PlaneProxyHandler.prototype.getOwnPropertyNames = function(){
  return Object.getOwnPropertyNames(this.obj);
};

PlaneProxyHandler.prototype.getPropertyNames = function(){ // !ES5
  return Object.getPropertyNames(this.obj);
};

PlaneProxyHandler.prototype.defineProperty = function(name, desc){
  return Object.defineProperty(this.obj, name, desc);
};

PlaneProxyHandler.prototype.delete = function(name){
  return delete this.obj[name];
};

PlaneProxyHandler.prototype.fix = function(){
  if(Object.isFrozen(this.obj)){
    return Object.getOwnPropertyNames(this.obj).map(function(name){
      return Object.getOwnPropertyDescriptor(this.obj, name);
    });
  }
  // As long as obj is not frozen, the proxy won't allow itself to be fixed
  return undefined; // will cause a TypeError to be thrown
};

// derived traps

PlaneProxyHandler.prototype.hasOwn = function(name){
  return Object.prototype.hasOwnProperty.call(this.obj, name);
};

PlaneProxyHandler.prototype.has = function(name){
  return name in this.obj;
};

PlaneProxyHandler.prototype.get = function(receiver, name){
/*
  var prop = this.obj[name];
  if(prop) return typeof prop === 'function' ? prop.bind(this.obj) : prop;
*/
  return this.obj[name];
};

PlaneProxyHandler.prototype.set = function(receiver, name, val){
  // 'set' needs checking property descriptor
  // ( bad behavior when set fails in non-strict mode )
//  console.log(' SET ' + name + ' VAL ' + val);
//  console.log(this.obj.__noSuchSetter__(name, val));
  if(this.obj[name]) this.obj[name] = val;
  else this.obj.__noSuchSetter__(name, val);
  return true;
};

PlaneProxyHandler.prototype.enumerate = function(){
  var result = [];
  for(var name in this.obj) result.push(name);
  return result;
};

PlaneProxyHandler.prototype.keys = function(){
  return Object.keys(this.obj);
};

var FunctionProxyHandler = function(){ // create instance at once (low memory)
  this.obj = function(){}; // override Object.getPrototypeOf(this).obj
//  this.tgt = targetobj; // this property must be set later (low memory)
//  this.calledMethod = method; // this property must be set later (low memory)
};

FunctionProxyHandler.prototype = new PlaneProxyHandler; // inherits obj
FunctionProxyHandler.prototype.tgt = (function(){})(); // undefined & overrided
FunctionProxyHandler.prototype.calledMethod = null; // null & overrided

FunctionProxyHandler.prototype.get = function(receiver, name){
  var fnc = this.obj; // reference is used in function(){} because other's this
  var tgt = this.tgt; // reference is used in function(){} because other's this
  var cm = this.calledMethod;
  // if(name == 'prototype') return tgt.prototype; // Function.prototype
  // if(name == 'constructor') ; // Proxy.createFunction(h, call, construct)
  if(name == 'inspect') return function(){ return tgt; } // util.inspect()
/*
  var prop = fnc[name];
  if(prop) return typeof prop === 'function' ? prop.bind(fnc) : prop;
*/
  if(fnc[name]) return fnc[name];
//  console.log(' get ' + name + ' ' + cm);
//  console.log(tgt.__noSuchGetter__(cm, []));
  // discussed about __noSuchProperty__
  // https://mail.mozilla.org/pipermail/es-discuss/2010-October/011930.html
  if(name == '_') return tgt.__noSuchGetter__(cm, []);
  return fnc[name] = function(){ return tgt.__noSuchGetter__(cm, []); };
};

var ObjectProxyHandler = function(target){ // each object must create instance
  this.obj = target; // override Object.getPrototypeOf(this).obj
};

ObjectProxyHandler.prototype = new PlaneProxyHandler; // inherits obj

ObjectProxyHandler.prototype.get = function(receiver, name){
  var obj = this.obj; // reference is used in function(){} because other's this
  // use encapsulated OCVariant in OLECall/OLEGet/OLESet accessor
  if(name == '__') return obj; // cyclic reference when set obj.__ as obj
  if(name == 'call'){
    return function(){
      var a = arguments.length < 2 ? [] : arguments[1];
      return obj.__noSuchMethod__(arguments[0], a); // ***
    };
  }
  if(name == 'get'){
    return function(){
      var a = arguments.length < 2 ? [] : arguments[1];
      return obj.__noSuchGetter__(arguments[0], a); // ***
    };
  }
  if(name == 'set'){
    return function(){
      var a = arguments.length < 2 ? [] : arguments[1];
      return obj.__noSuchSetter__(arguments[0], a); // ***
    };
  }
/*
  var prop = obj[name];
  if(prop) return typeof prop === 'function' ? prop.bind(obj) : prop;
*/
  if(obj[name]) return obj[name];
  fph.tgt = obj; // instance has been already existing (low memory)
  fph.calledMethod = name; // instance has been already existing (low memory)
  return Proxy.createFunction(fph, function(){
    return obj.__noSuchMethod__(
      name, Array.prototype.slice.call(arguments)); // ***
  });
};

var metamorphoses = function(obj){ // each object must create instance (HUGE)
  return Proxy.create(
    new ObjectProxyHandler(obj) // , Object.getPrototypeOf(obj)
  );
};

var fph = new FunctionProxyHandler; // create instance at once (low memory)

exports.metamorphoses = metamorphoses;
