var win32ole = require('win32ole');
var util = require('util');
win32ole.print('init_win32ole.test');
console.log(util.inspect(win32ole.verbose(), true, null, true));
var assert = require('assert');
var path = require('path');
var cwd = path.join(__dirname, '..');
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

/*
var xl = win32ole.client.Dispatch('Excel.Application');
xl.Visible = true;
var book = xl.Workbooks.Add();
var sheet = book.Worksheets(1);
sheet.Name = 'sheetnameA utf8';
sheet.Cells(1, 2).Value = 'test utf8';
var rg = sheet.Range(sheet.Cells(2, 2), sheet.Cells(4, 4));
rg.RowHeight = 5.18;
rg.ColumnWidth = 0.58;
rg.Interior.ColorIndex = 6; // Yellow
book.SaveAs('testfileutf8.xls');
xl.ScreenUpdating = true;
xl.Workbooks.Close();
xl.Quit();
*/

var st = new win32ole.Statement;
var xl = st.Dispatch('Excel.Application');
console.log(xl);
/*
xl.set('Visible', true);
var book = xl.get('Workbooks').call('Add', []);
var sheet = book.call('Worksheets', [1]);
sheet.set('Name', 'sheetnameA utf8');
sheet.call('Cells', [1, 2]).set('Value', 'test utf8');
var rg = sheet.call('Range',
  [sheet.call('Cells', [2, 2]), sheet.call('Cells', [4, 4])]);
rg.set('RowHeight', 5.18);
rg.set('ColumnWidth', 0.58);
rg.get('Interior').set('ColorIndex', 6); // Yellow
book.call('SaveAs', ['testfileutf8.xls']);
xl.set('ScreenUpdating', true);
xl.get('Workbooks').call('Close', []);
xl.call('Quit', []);
*/
st.Finalize(); // must be called now
