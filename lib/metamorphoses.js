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
  get: function(receiver, name){ return function(){ return obj[name]; }; },
  set: function(receiver, name, val){
    // 'set' needs checking property descriptor
    // ( bad behavior when set fails in non-strict mode )
//    console.log(' SET ' + name + ' VAL ' + val);
//    console.log(obj.__noSuchSetter__(name, val));
    obj[name] = val;
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
  var getsetHandler = planeProxyHandlerMaker({});
  getsetHandler.get = function(receiver, name){
//    console.log(' GET ' + calledMethod);
//    console.log(obj.__noSuchGetter__(calledMethod, []));
    return function(){ return 'Nothing'; };
  };
  var methodAccessHandler = planeProxyHandlerMaker(obj);
  methodAccessHandler.get = function(receiver, name){
    if(obj[name]) return obj[name];
    calledMethod = name;
    return Proxy.createFunction(getsetHandler, function(){
      return obj.__noSuchMethod__(name, Array.prototype.slice.call(arguments));
    });
  };
  return Proxy.create(methodAccessHandler);
};

exports.metamorphoses = metamorphoses;

var win32ole = require('win32ole');

try{
  var cl = metamorphoses(win32ole.client.Dispatch('OLEArgsTester.Server'));
  cl.__noSuchMethod__ = function(methodName, args){
    // console.log('method: ' + methodName + '(' + args + ')');
    return this.call(methodName, args);
  };
  cl.__noSuchGetter__ = function(getterName, args){
    // console.log('getter: ' + getterName + '(' + args + ')');
    return this.get(getterName, args);
  };
  cl.__noSuchSetter__ = function(setterName, args){
    // console.log('setter: ' + setterName + ' = ' + args[0]);
    return this.set(setterName, args[0]);
  };
  console.log('test');
//  console.log(cl.call('Test', ['a', 'b', 'c', 'd']).toUtf8());
  console.log(cl.Test('a1').toUtf8());
  console.log(cl.Test('a2', 'b2').toUtf8());
  console.log(cl.Test('a3', 'b3', 'c3').toUtf8());
  console.log(cl.Test('a4', 'b4', 'c4', 'd4').toUtf8());
  console.log('quit');
  cl.call('Quit');
  console.log('completed');
}catch(e){
  console.log('*** exception cached ***\n' + e);
}
