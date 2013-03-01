var win32ole = require('win32ole');
var util = require('util');
win32ole.print('init_win32ole.test\n');
console.log(util.inspect(win32ole.verbose(), true, null, true));
var assert = require('assert');
var path = require('path');
var cwd = path.join(win32ole.MODULEDIRNAME, '..');
assert.equal(__dirname, path.join(cwd, 'test'));
assert.equal(__filename, path.join(cwd, 'test/init_win32ole.test.js'));
assert.equal(require.resolve('win32ole'), path.join(cwd, 'lib/win32ole.js'));
console.log(win32ole.version());
console.log(win32ole.SOURCE_TIMESTAMP);
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

var fs = require('fs');
var tmpdir = path.join(cwd, 'test/tmp');
if(!fs.existsSync(tmpdir)) fs.mkdirSync(tmpdir);
var testfile = path.join(tmpdir, 'testfileutf8.xls');

var test_excel_ole = function(filename){
  // var xl = new ActiveXObject('Excel.Application'); // You may write it as:
  var xl = win32ole.client.Dispatch('Excel.Application');
  xl.Visible = true;
  var book = xl.Workbooks.Add();
  var sheet = book.Worksheets(1);
  try{
    sheet.Name = 'sheetnameA utf8';
    sheet.Cells(1, 2).Value = 'test utf8';
    var ltrb = [-4131, -4160, -4152, -4107]; // left, top, right, bottom
    for(var i = 0; i < ltrb.length; ++i){
      var bd = sheet.Cells(5, 6).Borders(ltrb[i]);
      bd.Weight = 2; // thin
      bd.LineStyle = 5; // dashdotdot
    }
    var rg = sheet.Range(sheet.Cells(2, 2), sheet.Cells(4, 4));
    rg.RowHeight = 5.18;
    rg.ColumnWidth = 0.58;
    rg.Interior.ColorIndex = 6; // Yellow
    console.log('saving to: "' + filename + '" ...');
    var result = book.SaveAs(filename);
    console.log(result);
  }catch(e){
    console.log('(exception cached)\n' + e);
  }
  xl.ScreenUpdating = true;
  xl.Workbooks.Close();
  xl.Quit();
};

try{
  test_excel_ole(testfile);
/*
  There are 3 ways to make force Garbage Collection for node.js / v8 .
  1. use huge memory to run GC automatically ( causes abnormal termination )
  2. win32ole.force_gc_extension(1);
  3. win32ole.force_gc_internal(1);
   ( see also examples/ole_args_test_client.js )
*/
}catch(e){
  console.log('*** exception cached ***\n' + e);
}
