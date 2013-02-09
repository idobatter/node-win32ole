var win32ole = require('win32ole');
win32ole.print('unicode.test start\n');
win32ole.printACP('森鷗外𠮟る\n'); // save utf8 only (not in cp932)
win32ole.print('unicode.test end\n');
