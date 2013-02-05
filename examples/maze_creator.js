var win32ole = require('win32ole');
win32ole.print('maze_creator\n');
console.log(win32ole.version());
var path = require('path');
var cwd = path.join(win32ole.MODULEDIRNAME, '..');

var fs = require('fs');
var tmpdir = path.join(cwd, 'test/tmp');
if(!fs.existsSync(tmpdir)) fs.mkdirSync(tmpdir);
var mazefile = path.join(tmpdir, 'maze_sample.xls');

// These parameters must be same as maze_solver.js
// colors 1: K 2: W 3: R 4: G 5: B 6: Y 7: M 8: C
var HEIGHT = 20, WIDTH = 30, OFFSET_ROW = 2, OFFSET_COL = 2;
var MAX_ROW = OFFSET_ROW + HEIGHT - 1, MAX_COL = OFFSET_COL + WIDTH - 1;
var sheet = null;

var mat = function(r, c){
  return sheet.get('Cells', [OFFSET_ROW + r, OFFSET_COL + c]);
};

var isPassed = function(r, c){
  try{
    return mat(r, c).get('Interior').get('ColorIndex').toInt32() != 6 // Y
  }catch(e){
    return true;
  }
};

var isDeadend = function(r, c){
  for(var d = 0; d < 4; ++d){
    var dr = d == 3 ? 1 : d == 2 ? -1 : 0;
    var dc = d == 1 ? 1 : d == 0 ? -1 : 0;
    if(!isPassed(r + dr, c + dc)) return false;
  }
  return true;
};

var drawWall = function(r, c, dlist){
  var e = mat(r, c);
  for(var d = 0; d < 4; ++d)
    if(dlist[d] == 0) e.get('Borders', [1 + d]).set('Weight', 2);
};

var dig = function(r, c, direc, count){
  var dlist = [0, 0, 0, 0];
  if(direc >= 0) dlist[[1, 0, 3, 2][direc]] = 1;
  mat(r, c).get('Interior').set('ColorIndex', 4); // G
  if(--count == 0) return drawWall(r, c, dlist);
  while(true){
    if(isDeadend(r, c)) return drawWall(r, c, dlist);
    var d = Math.floor(Math.random() * 4);
    var dr = d == 3 ? 1 : d == 2 ? -1 : 0;
    var dc = d == 1 ? 1 : d == 0 ? -1 : 0;
    if(!isPassed(r + dr, c + dc)){
      dlist[d] = 1;
      dig(r + dr, c + dc, d, count);
    }
  }
};

var maze_excel_ole = function(filename){
  var xl = win32ole.client.Dispatch('Excel.Application', 'C');
  xl.set('Visible', true);
  var book = xl.get('Workbooks').call('Add');
  // This code uses variable sheet as global
  sheet = book.get('Worksheets', [1]);
  try{
    sheet.set('Name', 'maze');
    sheet.get('Cells', [1, 2]).set('Value', 'maze');
    var rg = sheet.get('Range', [
      sheet.get('Cells', [OFFSET_ROW, OFFSET_COL]),
      sheet.get('Cells', [MAX_ROW, MAX_COL])]);
    rg.set('RowHeight', 5.18);
    rg.set('ColumnWidth', 0.58);
    rg.get('Interior').set('ColorIndex', 6); // Yellow
    // Math.random() seed is automatically set
    dig(HEIGHT - 1, WIDTH - 1, -1, WIDTH * HEIGHT);
    console.log('saving to: "' + filename + '" ...');
    var result = book.call('SaveAs', [filename]);
    console.log(result);
  }catch(e){
    console.log('(exception cached)\n' + e);
  }
  xl.set('ScreenUpdating', true);
  xl.get('Workbooks').call('Close');
  xl.call('Quit');
};

win32ole.client = new win32ole.Client;
try{
  maze_excel_ole(mazefile);
}catch(e){
  console.log('*** exception cached ***\n' + e);
}
win32ole.client.Finalize(); // must be called (version 0.0.x)
