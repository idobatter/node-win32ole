# NAME

node-win32ole - Asynchronous, non-blocking win32ole bindings for [node.js](https://github.com/joyent/node) .

win32ole makes accessibility from node.js to Excel, Word, Access, Outlook, InternetExplorer, WSH ( ActiveXObject ) and so on. It does not need TypeLibrary.


# USAGE

Install with `npm install win32ole`.

It works as... (version 0.1.x)

``` js
var win32ole = require('win32ole');
var xl = win32ole.client.Dispatch('Excel.Application', '.ACP'); // locale
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
```

But now it implements as... (version 0.0.x)

``` js
win32ole.client = new win32ole.Client;
try{
  var win32ole = require('win32ole');
  var xl = win32ole.client.Dispatch('Excel.Application', '.ACP'); // locale
  xl.set('Visible', true);
  var book = xl.get('Workbooks').call('Add');
  var sheet = book.get('Worksheets', [1]);
  try{
    sheet.set('Name', 'sheetnameA utf8');
    sheet.get('Cells', [1, 2]).set('Value', 'test utf8');
    var rg = sheet.get('Range',
      [sheet.get('Cells', [2, 2]), sheet.get('Cells', [4, 4])]);
    rg.set('RowHeight', 5.18);
    rg.set('ColumnWidth', 0.58);
    rg.get('Interior').set('ColorIndex', 6); // Yellow
    var result = book.call('SaveAs', ['testfileutf8.xls']);
    console.log(result);
  }catch(e){
    console.log('(exception cached)\n' + e);
  }
  xl.set('ScreenUpdating', true);
  xl.get('Workbooks').call('Close');
  xl.call('Quit');
}catch(e){
  console.log('*** exception cached ***\n' + e);
}
win32ole.client.Finalize(); // must be called (version 0.0.x)
```


# FEATURES

* So much implements.
* Implement accessors getter, setter and caller.
* npm


# API

See the [API documentation](https://github.com/idobatter/node-win32ole/wiki) in the wiki.


# BUILDING

This project uses VC++ 2008 Express (or later) and Python 2.6 (or later) .
(When using Python 2.5, it needs [multiprocessing 2.5 back port](http://pypi.python.org/pypi/multiprocessing/) .)

Bulding also requires node-gyp to be installed. You can do this with npm:

    npm install -g node-gyp

To obtain and build the bindings:

    git clone git://github.com/idobatter/node-win32ole.git
    cd node-win32ole
    node-gyp configure
    node-gyp build

You can also use [`npm`](https://github.com/isaacs/npm) to download and install them:

    npm install win32ole


# TESTS

[mocha](https://github.com/visionmedia/mocha) is required to run unit tests.

    npm install -g mocha
    nmake /a test


# CONTRIBUTORS

* [idobatter](https://github.com/idobatter)


# ACKNOWLEDGEMENTS

Inspired [Win32OLE](http://www.ruby-doc.org/stdlib/libdoc/win32ole/rdoc/)


# LICENSE

`node-win32ole` is [BSD licensed](https://github.com/idobatter/node-win32ole/raw/master/LICENSE).
