var win32ole = require('win32ole');
var util = require('util');
win32ole.print('init_win32ole.test');
console.log(util.inspect(win32ole.verbose(), true, null, true));
console.log(win32ole.version());
console.log(win32ole.SOURCE_TIMESTAMP);
var assert = require('assert');
var ref = require('ref');
var StructType = require('ref-struct');
var time_t = ref.types.long;
var suseconds_t = ref.types.long;
var timeval = StructType({
  tv_sec: time_t,
  tv_usec: suseconds_t
});
var tv = new timeval;
// console.log(util.inspect(tv.ref(), true, null, true));
assert.ok(win32ole.gettimeofday(tv.ref(), null));
console.log(tv.tv_sec + '.' + tv.tv_usec);
