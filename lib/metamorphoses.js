var Proxy = require('node-proxy');

var metamorphoses = function(obj){
  var planeProxyHandlerMaker = function(obj){
    return {
  // fundamental traps
  getOwnPropertyDescriptor: function(name){
    var desc = Object.getOwnPropertyDescriptor(obj, name);
    // a trapping proxy's properties must always be configurable
    if(desc !== undefined){ desc.configurable = true; }
    return desc;
  },
  getPropertyDescriptor: function(name){
    var desc = Object.getPropertyDescriptor(obj, name); // !ES5
    // a trapping proxy's properties must always be configurable
    if(desc !== undefined){ desc.configurable = true; }
    return desc;
  },
  getOwnPropertyNames: function(){ return Object.getOwnPropertyNames(obj); },
  getPropertyNames: function(){ return Object.getPropertyNames(obj); }, // !ES5
  defineProperty:
    function(name, desc){ return Object.defineProperty(obj, name, desc); },
  delete: function(name){ return delete obj[name]; },
  fix: function(){
    if(Object.isFrozen(obj)){
      return Object.getOwnPropertyNames(obj).map(function(name){
        return Object.getOwnPropertyDescriptor(obj, name);
      });
    }
    // As long as obj is not frozen, the proxy won't allow itself to be fixed
    return undefined; // will cause a TypeError to be thrown
  },
  // derived traps
  hasOwn:
    function(name){ return Object.prototype.hasOwnProperty.call(obj, name); },
  has: function(name){ return name in obj; },
  get: function(receiver, name){
/*
    var prop = obj[name];
    if(prop) return typeof prop === 'function' ? prop.bind(obj) : prop;
*/
    return obj[name];
  },
  set: function(receiver, name, val){
    // 'set' needs checking property descriptor
    // ( bad behavior when set fails in non-strict mode )
//    console.log(' SET ' + name + ' VAL ' + val);
//    console.log(obj.__noSuchSetter__(name, val));
    if(obj[name]) obj[name] = val;
    else obj.__noSuchSetter__(name, val);
    return true;
  },
  enumerate: function(){
    var result = [];
    for(var name in obj) result.push(name);
    return result;
  },
  keys: function(){ return Object.keys(obj); }
  // the end of new created planeProxyHandler
    };
  };

  var calledMethod;
  var func = function(){}; // undefined
  var getsetHandler = planeProxyHandlerMaker(func);
  getsetHandler.get = function(receiver, name){
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
    // use encapsulated OCVariant in OLECall/OLEGet/OLESet accessor
    if(name == '__') return obj;
    // discussed about __noSuchProperty__
    // https://mail.mozilla.org/pipermail/es-discuss/2010-October/011930.html
    if(name == '_') return obj.__noSuchGetter__(calledMethod, []);
    return func[name] = function(){
      return obj.__noSuchGetter__(calledMethod, []);
    };
  };
  var methodAccessHandler = planeProxyHandlerMaker(obj);
  methodAccessHandler.get = function(receiver, name){
    if(name == 'call'){
      return Proxy.createFunction(getsetHandler, function(){
        var a = arguments.length < 2 ? [] : arguments[1];
        return obj.__noSuchMethod__(arguments[0], a);
      });
    }
    if(name == 'get'){
      return Proxy.createFunction(getsetHandler, function(){
        var a = arguments.length < 2 ? [] : arguments[1];
        return obj.__noSuchGetter__(arguments[0], a);
      });
    }
    if(name == 'set'){
      return Proxy.createFunction(getsetHandler, function(){
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
    return Proxy.createFunction(getsetHandler, function(){
      return obj.__noSuchMethod__(name, Array.prototype.slice.call(arguments));
    });
  };
  return Proxy.create(methodAccessHandler); // , Object.getPrototypeOf(obj)
};

exports.metamorphoses = metamorphoses;
