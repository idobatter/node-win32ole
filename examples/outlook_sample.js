var win32ole = require('win32ole');
win32ole.print('outlook_sample\n');

var outlook_sample = function(){
  var ol = win32ole.client.Dispatch('Outlook.Application');
  // ol.Visible = true;
  var ns = ol.GetNameSpace('MAPI');
  var frcv = ns.GetDefaultFolder(6); // receive mail box tray
  var items = frcv.Items.valueOf(); // ***
  var count = items.Count;
  for(var n = 1; n <= count; ++n){
    win32ole.print(n + ' : ');
    var i = items.Item(n);
    win32ole.printACP(i.Subject);
    win32ole.print('\n');
  }
  // ol.Quit();
};

try{
  outlook_sample();
}catch(e){
  console.log('*** exception cached ***\n' + e);
}
