// This is a test program to create so much objects. (Cells = V8Variant)
// ( HEIGHT = 20, WIDTH = 30 )
var win32ole = require('win32ole');
win32ole.print('maze_solver\n');
console.log(win32ole.version());
var path = require('path');
var cwd = path.join(win32ole.MODULEDIRNAME, '..');

var fs = require('fs');
var tmpdir = path.join(cwd, 'test/tmp');
if(!fs.existsSync(tmpdir)) fs.mkdirSync(tmpdir);
var mazefile = path.join(tmpdir, 'maze_sample.xls');

// These parameters must be same as maze_creator.js
// colors 1: K 2: W 3: R 4: G 5: B 6: Y 7: M 8: C
var HEIGHT = 20, WIDTH = 30, OFFSET_ROW = 2, OFFSET_COL = 2;
var MAX_ROW = OFFSET_ROW + HEIGHT - 1, MAX_COL = OFFSET_COL + WIDTH - 1;
var sheet = null;

var mat = function(r, c){
  return sheet.Cells(OFFSET_ROW + r, OFFSET_COL + c);
};

var isExit = function(r, c){
  return (r == HEIGHT - 1) && (c == WIDTH - 1);
};

var isWall = function(r, c, d){
  var linestyle = mat(r, c).Borders(1 + d).LineStyle;
  if(linestyle == 1) return true;
  var dr = d == 3 ? 1 : d == 2 ? -1 : 0;
  var dc = d == 1 ? 1 : d == 0 ? -1 : 0;
  var color = mat(r + dr, c + dc).Interior.ColorIndex;
  if(color == 7 || color == 8) return true; // M or C
  return false;
};

var isDeadendWall = function(r, c, direc){
  if(direc < 0) return false;
  for(var d = 0; d < 4; ++d){
    if([1, 0, 3, 2][direc] == d) continue;
    if(!isWall(r, c, d)) return false;
  }
  return true;
};

var drawPath = function(r, c, solved, branch){
  var color = (solved && !branch) ? 8 : 7; // C or M
  mat(r, c).Interior.ColorIndex = color;
  return solved;
};

var dug = function(r, c, direc, solved, branch){
  var dlist = [0, 0, 0, 0];
  if(direc >= 0) dlist[[1, 0, 3, 2][direc]] = 1;
  mat(r, c).Interior.ColorIndex = 6; // Y
  while(true){
    if(isExit(r, c)) solved = true;
    if(isDeadendWall(r, c, direc)) return drawPath(r, c, solved, branch);
    var d = Math.floor(Math.random() * 4);
    if(dlist[d] == 1) continue;
    dlist[d] = 1;
    if(isWall(r, c, d)) continue;
    var dr = d == 3 ? 1 : d == 2 ? -1 : 0;
    var dc = d == 1 ? 1 : d == 0 ? -1 : 0;
    solved = dug(r + dr, c + dc, d, solved, solved);
  }
};

var solver_excel_ole = function(filename){
  var xl = win32ole.client.Dispatch('Excel.Application');
  xl.Visible = true;
  var book = xl.Workbooks.Open(filename);
  // This code uses variable sheet as global
  sheet = book.Worksheets(1);
  try{
    var rg = sheet.Range(
      sheet.Cells(OFFSET_ROW, OFFSET_COL), sheet.Cells(MAX_ROW, MAX_COL));
    // Math.random() seed is automatically set
    dug(0, 0, 3, false, false);
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
  solver_excel_ole(mazefile);
}catch(e){
  console.log('*** exception cached ***\n' + e);
}
