var win32ole = require('win32ole');
win32ole.print('outlook_sample\n');

var outlook_sample = function(){
  var ol = win32ole.client.Dispatch('Outlook.Application');
  // ol.set('Visible', true);
  var ns = ol.get('getNameSpace', ['MAPI']);
  var frcv = ns.get('GetDefaultFolder', [6]); // receive mail box tray
  var items = frcv.get('Items');
  var count = items.get('Count').toInt32();
  for(var n = 1; n <= count; ++n){
    win32ole.print(n + ' : ');
    var i = items.get('Item', [n]);
    win32ole.printACP(i.get('Subject').toUtf8());
    win32ole.print('\n');
  }
  // ol.call('Quit');
};

try{
  outlook_sample();
}catch(e){
  console.log('*** exception cached ***\n' + e);
}
