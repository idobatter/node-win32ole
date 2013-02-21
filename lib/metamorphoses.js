var Proxy = require('node-proxy');

var PlaneProxyHandler = function(target){
  this.obj = target;
};

PlaneProxyHandler.prototype.obj = null;

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

var metamorphoses = function(target){
  var obj = target;
  var calledMethod;
  var func = function(){}; // returns undefined
  var functionProxyHandler = new PlaneProxyHandler(func);
  functionProxyHandler.get = function(receiver, name){
    // if(name == 'prototype') return obj.prototype; // Function.prototype
    // if(name == 'constructor') ; // Proxy.createFunction(h, call, construct)
    if(name == 'inspect') return function(){ return obj; } // by util.inspect()
/*
    var prop = func[name];
    if(prop) return typeof prop === 'function' ? prop.bind(func) : prop;
*/
    if(func[name]) return func[name];
//    console.log(' get ' + name + ' ' + calledMethod);
//    console.log(obj.__noSuchGetter__(calledMethod, []));
    // discussed about __noSuchProperty__
    // https://mail.mozilla.org/pipermail/es-discuss/2010-October/011930.html
    if(name == '_') return obj.__noSuchGetter__(calledMethod, []);
    return func[name] = function(){
      return obj.__noSuchGetter__(calledMethod, []);
    };
  };
  var objectProxyHandler = new PlaneProxyHandler(obj);
  objectProxyHandler.get = function(receiver, name){
    // use encapsulated OCVariant in OLECall/OLEGet/OLESet accessor
    if(name == '__') return obj; // cyclic reference when set obj.__ as obj
    if(name == 'call'){
      return Proxy.createFunction(functionProxyHandler, function(){
        var a = arguments.length < 2 ? [] : arguments[1];
        return obj.__noSuchMethod__(arguments[0], a);
      });
    }
    if(name == 'get'){
      return Proxy.createFunction(functionProxyHandler, function(){
        var a = arguments.length < 2 ? [] : arguments[1];
        return obj.__noSuchGetter__(arguments[0], a);
      });
    }
    if(name == 'set'){
      return Proxy.createFunction(functionProxyHandler, function(){
        var a = arguments.length < 2 ? [] : arguments[1];
        return obj.__noSuchSetter__(arguments[0], a);
      });
    }
/*
    var prop = obj[name];
    if(prop) return typeof prop === 'function' ? prop.bind(obj) : prop;
*/
    if(obj[name]) return obj[name];
    calledMethod = name;
    return Proxy.createFunction(functionProxyHandler, function(){
      return obj.__noSuchMethod__(name, Array.prototype.slice.call(arguments));
    });
  };
  return Proxy.create(objectProxyHandler); // , Object.getPrototypeOf(obj)
};

exports.metamorphoses = metamorphoses;
