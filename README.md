# NAME

node-win32ole - Asynchronous, non-blocking win32ole bindings for [node.js](https://github.com/joyent/node) .


# USAGE

Install with `npm install node-win32ole`.

``` js
var win32ole = require('win32ole');
var xla = new win32ole.Dispatch('Excel.Application');

xla.series([
  function(){
    var wb = xla.Workbooks.Open('test/tmp/test.xls');
  },
  function(){
    var book = wb.Books(1);
  },
  function(){
    var ws = book.WorkSheets(1);
  },
  function(){
    book.Save();
    book.Close();
    xla.Quit();
  }
]);
```


# FEATURES

* So much implements.
* npm


# API

See the [API documentation](https://github.com/idobatter/node-win32ole/wiki) in the wiki.


# BUILDING

This project uses VC++ 2008 Express and Python 2.6 or later.

(When using Python 2.5, it needs multiprocessing 2.5 back port.)

Bulding also requires node-gyp to be installed. You can do this with npm:

    npm install -g node-gyp

To obtain and build the bindings:

    git clone git://github.com/idobatter/node-win32ole.git
    cd node-win32ole
    node-gyp configure
    node-gyp build

You can also use [`npm`](https://github.com/isaacs/npm) to download and install them:

    npm install node-win32ole


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
