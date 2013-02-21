var win32ole = require('win32ole');
win32ole.print('word_sample\n');
console.log(win32ole.version());
var path = require('path');
var cwd = path.join(win32ole.MODULEDIRNAME, '..');

var fs = require('fs');
var tmpdir = path.join(cwd, 'test/tmp');
if(!fs.existsSync(tmpdir)) fs.mkdirSync(tmpdir);
var outfile = path.join(tmpdir, 'word_sample.doc');

var word_sample = function(filename){
  var wd = win32ole.client.Dispatch('Word.Application');
  wd.Visible = true;
  var doc = wd.Documents._.Add(); // ***
  var para = doc.Content._.Paragraphs._.Add(); // ***
  para.Range._.Text = 'stringUTF8'; // ***
  try{
    console.log('saving to: "' + filename + '" ...');
    var result = doc.SaveAs(filename);
    console.log(result); // *** undefined
  }catch(e){
    console.log('(exception cached)\n' + e);
  }
  wd.Documents._.Close(); // ***
  wd.Quit();
};

try{
  word_sample(outfile);
}catch(e){
  console.log('*** exception cached ***\n' + e);
}
